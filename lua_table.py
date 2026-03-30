import re
from typing import Any


class LuaParseError(Exception):
    pass


class LuaTableParser:
    def __init__(self, text: str):
        self.text = text
        self.length = len(text)
        self.pos = 0

    def parse_assignment(self) -> tuple[str, Any]:
        self._skip_ws_and_comments()
        name = self._parse_identifier()
        self._skip_ws_and_comments()
        self._expect("=")
        self._skip_ws_and_comments()
        value = self._parse_value()
        self._skip_ws_and_comments()
        return name, value

    def _parse_value(self) -> Any:
        self._skip_ws_and_comments()
        ch = self._peek()
        if ch == "{":
            return self._parse_table()
        if ch == '"':
            return self._parse_string()
        if ch.isdigit() or ch in "-.":
            return self._parse_number()
        if self.text.startswith("true", self.pos):
            self.pos += 4
            return True
        if self.text.startswith("false", self.pos):
            self.pos += 5
            return False
        if self.text.startswith("nil", self.pos):
            self.pos += 3
            return None
        raise LuaParseError(f"Unexpected value at position {self.pos}")

    def _parse_table(self) -> dict[Any, Any]:
        table: dict[Any, Any] = {}
        self._expect("{")
        self._skip_ws_and_comments()

        while self.pos < self.length and self._peek() != "}":
            key = self._parse_key()
            self._skip_ws_and_comments()
            self._expect("=")
            self._skip_ws_and_comments()
            value = self._parse_value()
            table[key] = value
            self._skip_ws_and_comments()
            if self._peek() == ",":
                self.pos += 1
                self._skip_ws_and_comments()

        self._expect("}")
        return table

    def _parse_key(self) -> Any:
        self._skip_ws_and_comments()
        if self._peek() == "[":
            self._expect("[")
            self._skip_ws_and_comments()
            if self._peek() == '"':
                key = self._parse_string()
            else:
                key = self._parse_number()
            self._skip_ws_and_comments()
            self._expect("]")
            return key
        return self._parse_identifier()

    def _parse_identifier(self) -> str:
        self._skip_ws_and_comments()
        start = self.pos
        while self.pos < self.length and re.match(r"[A-Za-z0-9_]", self.text[self.pos]):
            self.pos += 1
        if start == self.pos:
            raise LuaParseError(f"Expected identifier at position {self.pos}")
        return self.text[start:self.pos]

    def _parse_string(self) -> str:
        self._expect('"')
        out = []
        while self.pos < self.length:
            ch = self.text[self.pos]
            self.pos += 1
            if ch == '"':
                return "".join(out)
            if ch == "\\":
                if self.pos >= self.length:
                    raise LuaParseError("Unterminated escape sequence")
                esc = self.text[self.pos]
                self.pos += 1
                if esc == "n":
                    out.append("\n")
                elif esc == "t":
                    out.append("\t")
                elif esc == "r":
                    out.append("\r")
                else:
                    out.append(esc)
            else:
                out.append(ch)
        raise LuaParseError("Unterminated string")

    def _parse_number(self) -> int | float:
        start = self.pos
        while self.pos < self.length and re.match(r"[0-9eE+\-\.]", self.text[self.pos]):
            self.pos += 1
        raw = self.text[start:self.pos]
        if not raw:
            raise LuaParseError(f"Expected number at position {start}")
        if "." in raw or "e" in raw.lower():
            return float(raw)
        return int(raw)

    def _skip_ws_and_comments(self) -> None:
        while self.pos < self.length:
            ch = self.text[self.pos]
            if ch.isspace():
                self.pos += 1
                continue
            if self.text.startswith("--", self.pos):
                self.pos += 2
                while self.pos < self.length and self.text[self.pos] not in "\r\n":
                    self.pos += 1
                continue
            break

    def _peek(self) -> str:
        if self.pos >= self.length:
            return ""
        return self.text[self.pos]

    def _expect(self, token: str) -> None:
        if not self.text.startswith(token, self.pos):
            raise LuaParseError(f"Expected '{token}' at position {self.pos}")
        self.pos += len(token)


def parse_lua_assignment(text: str) -> tuple[str, Any]:
    parser = LuaTableParser(text)
    return parser.parse_assignment()


def _escape_lua_string(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", "\\n")
        .replace("\t", "\\t")
        .replace("\r", "\\r")
    )


def _key_sorter(key: Any) -> tuple[int, Any]:
    if isinstance(key, int):
        return (0, key)
    return (1, str(key))


def _dump_value(value: Any, indent: int = 0) -> str:
    sp = "\t" * indent
    if isinstance(value, dict):
        lines = ["{"]
        for key in sorted(value.keys(), key=_key_sorter):
            if isinstance(key, int):
                key_repr = f"[{key}]"
            else:
                key_repr = f'["{_escape_lua_string(str(key))}"]'
            val_repr = _dump_value(value[key], indent + 1)
            lines.append(f"{sp}\t{key_repr} = {val_repr},")
        lines.append(f"{sp}" + "}")
        return "\n".join(lines)
    if isinstance(value, str):
        return f'"{_escape_lua_string(value)}"'
    if isinstance(value, bool):
        return "true" if value else "false"
    if value is None:
        return "nil"
    if isinstance(value, float):
        text = f"{value:.12f}".rstrip("0").rstrip(".")
        return text if text else "0"
    return str(value)


def dump_lua_assignment(name: str, value: Any) -> str:
    return f"{name} = \n{_dump_value(value, 0)}\n"
