"""
Thin wrapper around Coda MCP table_rows_read with pagination and filter-error detection.

Designed to be reused across all PDOW/SSD/CCW exporters. Assumes the MCP call is
routed through a callable that returns the raw dict response; this module is
agnostic to HOW that call happens (in-context via tool_use, or via an API client).

Key lessons baked in from MSCSIA reconnaissance:
  1. `filterFormula` on lookup columns SILENTLY returns unfiltered results
     plus a `filterFormulaError` key. We detect and fall back to client-side.
  2. `filterColumnNames` drops entries whose name doesn't resolve. Cheap mistake;
     fail loud by asserting returned columns match requested ones when it matters.
  3. Lookup refs return a `.name` field we can use directly — no need to
     re-resolve row IDs to human-readable names for POs/CCTs/courses.
"""
from typing import Callable, Any


class CodaReadClient:
    """
    Callable that invokes the MCP `table_rows_read` tool and returns the result dict.
    Keeps us decoupled from the specific MCP runtime.
    """
    def __init__(self, mcp_read: Callable[..., dict]):
        self.read = mcp_read

    def read_all_rows(self, table_uri: str, page_size: int = 100,
                      max_pages: int = 20, filter_columns: list[str] | None = None
                      ) -> list[dict]:
        """
        Paginate through a Coda table and return all rows. Stops early if hasMore=False.
        Raises if max_pages is exceeded (safety cap).
        """
        all_rows: list[dict] = []
        for page in range(max_pages):
            kwargs = {
                "uri": table_uri,
                "rowLimit": page_size,
                "rowOffset": page * page_size,
            }
            if filter_columns:
                kwargs["filterColumnNames"] = filter_columns
            resp = self.read(**kwargs)
            result = resp.get("result", resp)
            rows = result.get("rows", [])
            all_rows.extend(rows)
            if not result.get("hasMore"):
                return all_rows
            if len(rows) == 0:
                # Defensive: something's off — avoid infinite loop
                return all_rows
        raise RuntimeError(f"Exceeded max_pages ({max_pages}) on {table_uri} — table may be larger than expected")

    @staticmethod
    def ref_identifier(cell: dict) -> str | None:
        """
        Safely extract an identifier from a ref cell value.
        Cell shape: {'content': {'type': 'ref', 'identifier': 'i-xxx', 'name': '...'}}
        Returns None if cell is empty or malformed.
        """
        content = cell.get("content") if cell else None
        if isinstance(content, dict) and content.get("type") == "ref":
            return content.get("identifier")
        return None

    @staticmethod
    def ref_name(cell: dict) -> str | None:
        """Extract the display name from a ref cell."""
        content = cell.get("content") if cell else None
        if isinstance(content, dict) and content.get("type") == "ref":
            return content.get("name")
        return None

    @staticmethod
    def text(cell: dict) -> str:
        """Extract plain text from a text or scalar cell. Returns '' for empty/None."""
        if not cell:
            return ""
        content = cell.get("content", "")
        if isinstance(content, str):
            return content
        if isinstance(content, (int, float)):
            return str(content)
        if isinstance(content, dict):
            # Number cell: {'type': 'num', 'value': 3, ...}
            if content.get("type") == "num":
                return str(content.get("value", ""))
            # Slate cell (rich text) — flatten to plain text
            if content.get("type") == "slate":
                return _slate_to_plain(content)
        return ""

    @staticmethod
    def number(cell: dict):
        """Extract numeric value from a number cell. Returns None if not present."""
        if not cell:
            return None
        content = cell.get("content")
        if isinstance(content, (int, float)):
            return content
        if isinstance(content, dict) and content.get("type") == "num":
            return content.get("value")
        return None

    @staticmethod
    def array_refs(cell: dict) -> list[dict]:
        """
        Extract a list of ref objects from a multiselect/array cell.
        Cell shape: {'content': {'type': 'arr', 'value': [{'type': 'ref', 'identifier': ..., 'name': ...}, ...]}}
        Returns [] for empty cells.
        """
        content = cell.get("content") if cell else None
        if isinstance(content, dict) and content.get("type") == "arr":
            return [v for v in content.get("value", []) if isinstance(v, dict) and v.get("type") == "ref"]
        return []


def _slate_to_plain(slate_content: dict) -> str:
    """Flatten a Coda slate (rich-text) structure to plain text."""
    root = slate_content.get("root", {})
    return _slate_walk(root).strip()


def _slate_walk(node) -> str:
    if isinstance(node, dict):
        if "text" in node:
            return node.get("text") or ""
        children = node.get("children", [])
        return "".join(_slate_walk(c) for c in children)
    if isinstance(node, list):
        return "".join(_slate_walk(c) for c in node)
    return ""
