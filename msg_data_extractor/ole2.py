import struct


class OLE2Reader:
    """Pure-Python parser for OLE2 / Compound File Binary format (.msg)."""

    MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF
    MAX_VALID = 0xFFFFFFFA

    def __init__(self, path: str):
        with open(path, "rb") as fh:
            self.data = fh.read()
        if self.data[:8] != self.MAGIC:
            raise ValueError(f"Not an OLE2 file: {path}")
        self._parse_header()
        self._build_fat()
        self._read_directory()

    def _parse_header(self):
        data = self.data
        self.sector_size = 1 << struct.unpack_from("<H", data, 0x1E)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 0x20)[0]
        self.first_dir_sector = struct.unpack_from("<I", data, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", data, 0x38)[0]
        self.first_minifat_sec = struct.unpack_from("<I", data, 0x3C)[0]
        self.difat = list(struct.unpack_from("<109I", data, 0x4C))

    def _sector_data(self, sector: int) -> bytes:
        offset = (sector + 1) * self.sector_size
        return self.data[offset : offset + self.sector_size]

    def _chain(self, start: int) -> list[int]:
        sectors = []
        current = start
        seen = set()
        while current <= self.MAX_VALID:
            if current in seen:
                break
            seen.add(current)
            sectors.append(current)
            if current >= len(self.fat):
                break
            current = self.fat[current]
        return sectors

    def _read_chain(self, start: int, size: int) -> bytes:
        raw = b"".join(self._sector_data(sector) for sector in self._chain(start))
        return raw[:size]

    def _build_fat(self):
        fat_bytes = b""
        for sector in self.difat:
            if sector > self.MAX_VALID:
                break
            fat_bytes += self._sector_data(sector)
        count = len(fat_bytes) // 4
        self.fat = list(struct.unpack_from(f"<{count}I", fat_bytes))

    def _build_minifat(self):
        if self.first_minifat_sec > self.MAX_VALID:
            self.minifat = []
            return
        fat_bytes = b"".join(
            self._sector_data(sector)
            for sector in self._chain(self.first_minifat_sec)
        )
        count = len(fat_bytes) // 4
        self.minifat = list(struct.unpack_from(f"<{count}I", fat_bytes))

    def _read_mini_chain(self, start: int, size: int) -> bytes:
        mini_stream = b"".join(
            self._sector_data(sector)
            for sector in self._chain(self.root["start"])
        )
        sectors = []
        current = start
        seen = set()
        while current <= self.MAX_VALID:
            if current in seen:
                break
            seen.add(current)
            sectors.append(current)
            if current >= len(self.minifat):
                break
            current = self.minifat[current]
        raw = b"".join(
            mini_stream[
                sector * self.mini_sector_size : (sector + 1) * self.mini_sector_size
            ]
            for sector in sectors
        )
        return raw[:size]

    def _read_directory(self):
        dir_bytes = b"".join(
            self._sector_data(sector)
            for sector in self._chain(self.first_dir_sector)
        )
        self.entries = []
        for index in range(len(dir_bytes) // 128):
            raw = dir_bytes[index * 128 : (index + 1) * 128]
            name_len = struct.unpack_from("<H", raw, 64)[0]
            name = (
                raw[: name_len - 2].decode("utf-16-le", errors="ignore")
                if name_len > 2
                else ""
            )
            obj_type = raw[66]
            left, right, child = [
                struct.unpack_from("<I", raw, pos)[0] for pos in [68, 72, 76]
            ]
            start, size = struct.unpack_from("<II", raw, 116)
            self.entries.append(
                {
                    "name": name,
                    "type": obj_type,
                    "left": left,
                    "right": right,
                    "child": child,
                    "start": start,
                    "size": size,
                }
            )
        self.root = self.entries[0]
        self._build_minifat()

    def read_stream(self, index: int) -> bytes:
        entry = self.entries[index]
        if entry["size"] < self.mini_cutoff:
            return self._read_mini_chain(entry["start"], entry["size"])
        return self._read_chain(entry["start"], entry["size"])

    def _subtree(self, index: int) -> list[int]:
        if index > self.MAX_VALID or index >= len(self.entries):
            return []
        entry = self.entries[index]
        return [index] + self._subtree(entry["left"]) + self._subtree(entry["right"])

    def children_of(self, storage_idx: int) -> list[int]:
        child = self.entries[storage_idx]["child"]
        return self._subtree(child) if child <= self.MAX_VALID else []

    def find_in(self, storage_idx: int, name: str) -> int | None:
        for index in self.children_of(storage_idx):
            if self.entries[index]["name"] == name:
                return index
        return None

    def get_string(self, storage_idx: int, prop: str) -> str | None:
        index = self.find_in(storage_idx, prop)
        if index is None:
            return None
        raw = self.read_stream(index)
        return raw.decode("utf-16-le", errors="ignore").rstrip("\x00")

    def get_bytes(self, storage_idx: int, prop: str) -> bytes | None:
        index = self.find_in(storage_idx, prop)
        if index is None:
            return None
        return self.read_stream(index)

    def attachment_storages(self) -> list[int]:
        return [
            index
            for index in self.children_of(0)
            if self.entries[index]["name"].startswith("__attach_version1.0_#")
        ]
