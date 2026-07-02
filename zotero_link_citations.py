"""
zotero_link_citations.py
────────────────────────
Replaces manual citations in a .docx with native Zotero fields
(ADDIN ZOTERO_ITEM), searching references in a local Zotero SQLite database.

Dependencies:
    pip install python-docx lxml
"""

import json
import re
import sqlite3
import sys
import unicodedata
from copy import deepcopy
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Missing dependencies. Install with:  pip install python-docx lxml")
    sys.exit(1)


# ─── Normalization ────────────────────────────────────────────────────────

def normalize(text: str) -> str:
    text = unicodedata.normalize("NFD", text.lower())
    return "".join(c for c in text if unicodedata.category(c) != "Mn")


# ─── Citation parsing ─────────────────────────────────────────────────────
#
# Supported author forms (NAME can contain apostrophes, hyphens, accents):
#   Single name     : "Nom", "O'Brien", "Smith-Jones", "NOAA"
#   Two names       : "Pre Nom", "P. Nom", "P. R. Nom"
#   Trailing init.  : "Nom P.", "Nom P. R."
#   et al.          : "Nom et al.", "Nom et al"
#   Two authors     : "Nom1 and Nom2"
#   Same-auth years : "Chan et al., 2019, 2024"
# Year: 1700-2099 with optional lowercase suffix (2019a).

_YEAR_PAT = r"(?:1[7-9]|20)\d{2}[a-z]?"

# One capitalised name token (letters, apostrophes, hyphens, accented chars)
# May end with a dot to handle initials like "X."
_N  = r"[A-ZÀ-Ö][A-ZÀ-ÿa-z'\-]*\.?"
# A name-or-initial token
_NI = _N
# Lowercase particles that may appear between name tokens (von, van, de, del, …)
_PARTICLE = r"(?:von|van|de|del|di|le|la|du|den|der|ten|ter|al)\s+"

_CITE_TOKEN_RE = re.compile(
    # ── author ─────────────────────────────────────────────────────────────
    # Greedily consume capitalised tokens (with optional lowercase particles),
    # stopping before "et al.", "and NAME", or a year.
    r"(?P<author>"
        r"(?:" + _PARTICLE + r")*"              # optional leading particle(s)
        r"(?:" + _NI + r")"                     # first cap token (required)
        r"(?:"
            r"(?!\s+et\s+al)"
            r"(?!\s+and\s+[A-ZÀ-Öa-z])"
            r"(?!\s*[,\s]+(?:1[7-9]|20)\d{2})"
            r"\s+(?:" + _PARTICLE + r")*"       # optional particle before next token
            r"(?:" + _NI + r")"
        r")*"
    r")"
    # ── optional et al. / and NAME2 ────────────────────────────────────────
    r"(?:"
        r"\s+et\s+al\.?"
        r"|"
        r"\s+and\s+(?P<author2>"
            r"(?:" + _PARTICLE + r")*"
            r"(?:" + _NI + r")"
            r"(?:"
                r"(?!\s+et\s+al)(?!\s*[,\s]+(?:1[7-9]|20)\d{2})"
                r"\s+(?:" + _PARTICLE + r")*(?:" + _NI + r")"
            r")*"
        r")"
    r")?"
    # ── separator + primary year ────────────────────────────────────────────
    r"[\s,]+"
    r"(?P<year>" + _YEAR_PAT + r")"
    # ── optional same-author extra years ───────────────────────────────────
    r"(?P<extra>(?:\s*,\s*" + _YEAR_PAT + r")*)",
    re.UNICODE,
)

# Block detection: parenthesised content with at least one year
CITATION_BLOCK_RE = re.compile(
    r"\(([^()]*(?:" + _YEAR_PAT + r")[^()]*)\)"
)

_HAS_YEAR_RE = re.compile(_YEAR_PAT)


def _tokenize_block(block_text: str) -> tuple[list[dict], str, str]:
    """
    Parse text between parentheses → (refs, clean_inner, leftover_suffix).

    Each ref dict now includes a 'matched_text' key with the exact
    text that was matched for that citation (used to build clean display text).
    """
    refs           = []
    clean_parts    = []
    leftover_parts = []
    pos            = 0
    text           = block_text

    while pos < len(text):
        m = _CITE_TOKEN_RE.search(text, pos)
        if not m:
            tail = text[pos:].strip().strip(",;").strip()
            if tail:
                leftover_parts.append(tail)
            break

        before = text[pos:m.start()].strip().strip(",;").strip()
        if before:
            leftover_parts.append(before)

        author      = m.group("author").strip()
        author2     = (m.group("author2") or "").strip()  # second author if "and NAME2"
        year_raw    = m.group("year")
        matched_txt = m.group(0).strip()

        refs.append({"author": author, "author2": author2, "year": year_raw[:4],
                     "year_suffix": year_raw[4:], "matched_text": matched_txt})
        clean_parts.append(matched_txt)

        # Same-author extra years captured in 'extra' group
        extra = m.group("extra") or ""
        for yr in re.findall(_YEAR_PAT, extra):
            refs.append({"author": author, "author2": author2, "year": yr[:4],
                         "year_suffix": yr[4:], "matched_text": matched_txt})

        pos = m.end()

    clean_inner   = "; ".join(p for p in clean_parts if p)
    leftover_text = ", ".join(p for p in leftover_parts if p)
    return refs, clean_inner, leftover_text


def find_citation_blocks(text: str) -> list[tuple]:
    """
    Returns list of (full_matched_text, refs, clean_inner, leftover_suffix).
    Only blocks with at least one valid citation are returned.
    Blocks containing 'REF: NOT FOUND' or 'REF: MULTI CHOICE' are skipped —
    they are error markers inserted by a previous run of this script.
    """
    results = []
    for m in CITATION_BLOCK_RE.finditer(text):
        inner = m.group(1)
        # Skip blocks that are already-processed error annotations
        if "REF: NOT FOUND" in inner or "REF: MULTI CHOICE" in inner:
            continue
        refs, clean_inner, leftover = _tokenize_block(inner)
        if refs:
            results.append((m.group(0), refs, clean_inner, leftover))
    return results

# ─── Zotero SQLite database ───────────────────────────────────────────────

class ZoteroDB:
    def __init__(self, db_path: str, library_id: int | None = None):
        path = Path(db_path).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"Zotero database not found: {path}")
        try:
            self.conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
            self.conn.row_factory = sqlite3.Row
            # Test the connection immediately to catch a locked DB early
            self.conn.execute("SELECT 1")
        except sqlite3.OperationalError as e:
            print(f"\nError: cannot open Zotero database ({e})")
            print("→ Please close the Zotero app and try again.")
            sys.exit(1)
        self.library_id = library_id
        self._cache: dict = {}
        print(f"  Zotero DB: {path}")
        if library_id:
            print(f"  Library ID: {library_id}")

    @staticmethod
    def list_libraries(db_path: str) -> list[dict]:
        path = Path(db_path).expanduser()
        try:
            conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
            conn.row_factory = sqlite3.Row
            rows = conn.execute("""
                SELECT l.libraryID, l.type,
                       COALESCE(g.name, 'Personal Library') AS name
                FROM libraries l
                LEFT JOIN groups g ON g.libraryID = l.libraryID
                ORDER BY l.libraryID
            """).fetchall()
            conn.close()
        except sqlite3.OperationalError as e:
            print(f"\nError: cannot read Zotero database ({e})")
            print("→ Please close the Zotero app and try again.")
            sys.exit(1)
        return [dict(r) for r in rows]

    def search(self, author: str, year: str, author2: str = "") -> list[dict]:
        # Strip trailing initials like "X." or "X. Y." from the author string.
        # Zotero stores only the last name, so "Truc X." should match "Truc".
        def _strip_initials(name: str) -> str:
            # Remove trailing space+initial(s) like " X." or " X. Y."
            return re.sub(r'(\s+[A-ZÀ-Ö]\.?)+$', '', name).strip()

        author_clean  = _strip_initials(author)
        author2_clean = _strip_initials(author2) if author2 else ""
        key = f"{normalize(author_clean)}|{year}|{normalize(author2_clean)}"
        if key in self._cache:
            return self._cache[key]

        lib_filter = "AND i.libraryID = ?" if self.library_id else ""
        params = [f"%{year}%"]
        if self.library_id:
            params.append(self.library_id)

        query = f"""
            SELECT DISTINCT i.itemID, i.key AS zotero_key, i.libraryID,
                   it.typeName AS itemType
            FROM items i
            JOIN itemTypes it ON it.itemTypeID = i.itemTypeID
            JOIN itemData id_date ON id_date.itemID = i.itemID
                AND id_date.fieldID IN (
                    SELECT fieldID FROM fields WHERE fieldName IN ('date','year'))
            JOIN itemDataValues idv_date ON idv_date.valueID = id_date.valueID
                AND idv_date.value LIKE ?
            WHERE i.itemID NOT IN (SELECT itemID FROM deletedItems)
              AND it.typeName NOT IN ('attachment','note')
              {lib_filter}
        """
        try:
            rows = self.conn.execute(query, params).fetchall()
        except sqlite3.OperationalError as e:
            print(f"\nError: Zotero database became locked during scan ({e})")
            print("→ Please close the Zotero app and try again.")
            sys.exit(1)

        matches = []
        for row in rows:
            creators = self.get_all_creators(row["itemID"])
            if not creators:
                continue
            first_last = creators[0].get("lastName", "")
            norm_zotero = normalize(first_last)
            norm_author = normalize(author_clean)

            # A match requires that one name is a prefix of the other AND
            # the prefix ends on a word boundary (not mid-word).
            # e.g. "ansari" matches "ansari" or "ansari-smith" but NOT "an"
            # e.g. "richardson" matches "richardson" or "rich" (truncated) but NOT "ri"
            def prefix_match(shorter: str, longer: str) -> bool:
                if not longer.startswith(shorter):
                    return False
                if len(shorter) == len(longer):
                    return True  # exact match
                # Next char in longer must be a non-letter (hyphen, space, end)
                next_char = longer[len(shorter)]
                return not next_char.isalpha()

            match = (
                prefix_match(norm_author, norm_zotero)
                or prefix_match(norm_zotero, norm_author)
            )

            if match:
                # If a second author was specified (e.g. "Li and Paul, 2026"),
                # verify it against the second creator in Zotero
                if author2_clean:
                    if len(creators) < 2:
                        match = False
                    else:
                        second_last = creators[1].get("lastName", "")
                        norm_a2       = normalize(author2_clean)
                        norm_second   = normalize(second_last)
                        match = (
                            prefix_match(norm_a2, norm_second)
                            or prefix_match(norm_second, norm_a2)
                        )

            if match:
                fields = self.get_item_fields(row["itemID"])
                matches.append({**dict(row), "fields": fields, "creators": creators})

        # Deduplicate by title: if multiple Zotero items share the same title,
        # keep only the first (they are duplicates in Zotero, not true homonyms).
        seen_titles: dict[str, int] = {}  # normalized title → index of first occurrence
        deduped = []
        for item in matches:
            title = normalize(item["fields"].get("title", ""))
            if title not in seen_titles:
                seen_titles[title] = len(deduped)
                deduped.append(item)
            # else: duplicate title → silently skip

        self._cache[key] = deduped
        return deduped

    def get_all_creators(self, item_id: int) -> list[dict]:
        rows = self.conn.execute("""
            SELECT cr.lastName, cr.firstName, ic.orderIndex,
                   ct.creatorType AS creatorType
            FROM itemCreators ic
            JOIN creators cr ON cr.creatorID = ic.creatorID
            JOIN creatorTypes ct ON ct.creatorTypeID = ic.creatorTypeID
            WHERE ic.itemID = ?
            ORDER BY ic.orderIndex
        """, (item_id,)).fetchall()
        return [dict(r) for r in rows]

    def get_item_fields(self, item_id: int) -> dict:
        rows = self.conn.execute("""
            SELECT f.fieldName, idv.value
            FROM itemData id
            JOIN fields f ON f.fieldID = id.fieldID
            JOIN itemDataValues idv ON idv.valueID = id.valueID
            WHERE id.itemID = ?
        """, (item_id,)).fetchall()
        return {r["fieldName"]: r["value"] for r in rows}

    def close(self):
        self.conn.close()


# ─── Zotero field builder ─────────────────────────────────────────────────

def _item_data_for(item: dict) -> dict:
    """Build the CSL item_data dict for one resolved Zotero item.
    Follows the exact structure Zotero expects:
    - authors as {"family": ..., "given": ...}
    - issued as {"date-parts": [["YYYY"]]}  (year as string)
    """
    creators  = item["creators"]
    fields    = item["fields"]
    year_str  = fields.get("date", fields.get("year", ""))
    year_val  = year_str[:4] if year_str and year_str[:4].isdigit() else ""

    # Build author list using CSL keys (family/given, not lastName/firstName)
    csl_authors = []
    for c in creators:
        last  = c.get("lastName", "").strip()
        first = c.get("firstName", "").strip()
        if last and first:
            csl_authors.append({"family": last, "given": first})
        elif last:
            csl_authors.append({"family": last})
        elif first:
            # Some entries store full name in firstName only
            csl_authors.append({"literal": first})

    item_data = {
        "id":     item["zotero_key"],   # use the real Zotero key as id
        "type":   item.get("itemType", "article-journal"),
        "title":  fields.get("title", ""),
        "author": csl_authors,
        # date-parts must be a list of lists, with year as a STRING
        "issued": {"date-parts": [[year_val]] if year_val else [[]]},
    }
    for zf, csljf in [("publicationTitle", "container-title"), ("volume", "volume"),
                       ("issue", "issue"), ("pages", "page"), ("DOI", "DOI"), ("url", "URL")]:
        if zf in fields:
            item_data[csljf] = fields[zf]
    return item_data


def build_zotero_field_xml(items: list[dict], display_text: str) -> str:
    """
    Build the ADDIN ZOTERO_ITEM field instruction for one or more resolved items.
    Matches the exact JSON structure that Zotero's Word plugin expects.
    """
    citation_items = []
    for item in items:
        idata = _item_data_for(item)
        # URI must use the library-specific path
        lib_id = item.get("libraryID", 1)
        key    = item["zotero_key"]
        if lib_id == 1:
            uri = f"http://zotero.org/users/local/{key}"
        else:
            uri = f"http://zotero.org/groups/{lib_id}/items/{key}"
        citation_items.append({
            "id":       key,
            "uris":     [uri],
            "itemData": idata,
        })

    citation_id = f"cite-{items[0]['zotero_key']}"
    citation_data = {
        "citationID":    citation_id,
        "properties":    {
            "formattedCitation": display_text,
            "plainCitation":     display_text,
            "noteIndex":         0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"
    }
    return f'ADDIN ZOTERO_ITEM CSL_CITATION {json.dumps(citation_data, ensure_ascii=False)}'


# ─── Word XML helpers ─────────────────────────────────────────────────────

def _count_zotero_fields(para) -> int:
    """Count existing ZOTERO_ITEM fields in the paragraph."""
    count = 0
    for instr in para._p.iter(qn("w:instrText")):
        if (instr.text or "").strip().startswith("ADDIN ZOTERO_ITEM"):
            count += 1
    return count


def _get_run_text(run) -> str:
    t = run.find(qn("w:t"))
    return (t.text or "") if t is not None else ""


def _set_run_text(run, text: str):
    t = run.find(qn("w:t"))
    if t is None:
        t = OxmlElement("w:t")
        run.append(t)
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text


def _get_plain_runs(para) -> list:
    """
    Return plain w:r elements, skipping field runs and runs inside fields.
    Crucially, also skips the display-text run of comment-anchored fields —
    those have a w:rPr/w:rStyle with val='CommentReference' or contain
    w:commentReference, but more reliably: we just track fldChar depth so
    any run inside a field (including comment reference fields) is excluded.
    Non-run elements (commentRangeStart/End, bookmarkStart/End, …) are
    never iterated here — they simply won't appear in the result.
    """
    plain = []
    in_field = 0
    for child in para._p:
        if child.tag != qn("w:r"):
            continue
        fc = child.find(qn("w:fldChar"))
        if fc is not None:
            ft = fc.get(qn("w:fldCharType"), "")
            if ft == "begin":
                in_field += 1
            elif ft == "end":
                in_field = max(0, in_field - 1)
            continue
        if child.find(qn("w:instrText")) is not None:
            continue
        if in_field > 0:
            continue
        # Skip runs that contain only a comment reference mark
        if child.find(qn("w:commentReference")) is not None:
            continue
        plain.append(child)
    return plain


def _para_plain_text(para) -> str:
    """Full plain text of a paragraph (for display in interactive prompts)."""
    return "".join(_get_run_text(r) for r in _get_plain_runs(para))


def _split_run_at(run, offset: int):
    """
    Split run at offset. run keeps text[:offset], a clone gets text[offset:].
    Returns (left_run, right_run).
    """
    text = _get_run_text(run)
    _set_run_text(run, text[:offset])
    right = deepcopy(run)
    _set_run_text(right, text[offset:])
    p = run.getparent()
    p.insert(list(p).index(run) + 1, right)
    return run, right


def _make_field_elements(field_instr: str, display_text: str, ref_rpr) -> list:
    def new_run(child):
        r = OxmlElement("w:r")
        if ref_rpr is not None:
            r.append(deepcopy(ref_rpr))
        r.append(child)
        return r

    fc_begin = OxmlElement("w:fldChar")
    fc_begin.set(qn("w:fldCharType"), "begin")
    instr_elem = OxmlElement("w:instrText")
    instr_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr_elem.text = f" {field_instr} "
    fc_sep = OxmlElement("w:fldChar")
    fc_sep.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = display_text
    r_display = OxmlElement("w:r")
    if ref_rpr is not None:
        r_display.append(deepcopy(ref_rpr))
    r_display.append(t)
    fc_end = OxmlElement("w:fldChar")
    fc_end.set(qn("w:fldCharType"), "end")
    return [new_run(fc_begin), new_run(instr_elem), new_run(fc_sep), r_display, new_run(fc_end)]


def _set_red_text(run):
    """Set run text color to red (works in Word 2019 and all versions)."""
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run.insert(0, rpr)
    for c in rpr.findall(qn("w:color")):
        rpr.remove(c)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "FF0000")
    rpr.append(color)


def _make_plain_run(text: str, rpr) -> object:
    """Create a plain w:r with the given text, inheriting rpr formatting."""
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(deepcopy(rpr))
    t_elem = OxmlElement("w:t")
    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_elem.text = text
    r.append(t_elem)
    return r


# ─── Surgical paragraph replacement ───────────────────────────────────────

def _get_run_span(p_elem, plain_runs: list, text_start: int, text_end: int) -> tuple[int, int]:
    """
    Given character positions [text_start, text_end) in the plain text,
    return the [child_idx_start, child_idx_end] in p_elem's all_children.
    """
    all_children = list(p_elem)
    pos = 0
    first_run = None
    last_run  = None
    for r in plain_runs:
        t = _get_run_text(r)
        run_start = pos
        run_end   = pos + len(t)
        if run_end > text_start and run_start < text_end:
            if first_run is None:
                first_run = r
            last_run = r
        pos = run_end
    if first_run is None:
        return (-1, -1)
    try:
        return (all_children.index(first_run), all_children.index(last_run))
    except ValueError:
        return (-1, -1)


def _attribute_comments_to_spans(
    p_elem,
    span_comments: list,
    linked_span: tuple[int, int],    # (i_start, i_end) of linked runs
    red_span:    tuple[int, int],     # (i_start, i_end) of red runs
) -> tuple[list, list]:
    """
    Attribute each comment to either the linked or the red span,
    based on which span it overlaps more.
    Returns (linked_comments, red_comments) where each is a list of
    (cid, crs, cre, crr) tuples.
    """
    linked_comments = []
    red_comments    = []

    def overlap(a_start, a_end, b_start, b_end) -> int:
        return max(0, min(a_end, b_end) - max(a_start, b_start) + 1)

    for cid, crs, cre, crr in span_comments:
        # Find the child indices of this comment's start and end
        all_children = list(p_elem)
        cs = all_children.index(crs) if crs is not None and crs in all_children else -1
        ce = all_children.index(cre) if cre is not None and cre in all_children else cs

        linked_overlap = overlap(cs, ce, linked_span[0], linked_span[1])
        red_overlap    = overlap(cs, ce, red_span[0],    red_span[1])

        if linked_overlap >= red_overlap and linked_overlap > 0:
            linked_comments.append((cid, crs, cre, crr))
        else:
            red_comments.append((cid, crs, cre, crr))

    return linked_comments, red_comments


def _reinsert_comments(p_elem, comments: list, before_elem, after_elem) -> None:
    """
    Insert comment elements: all crs before before_elem, all cre+crr after after_elem.
    """
    if not comments:
        return
    all_ch = list(p_elem)
    if before_elem is None or before_elem not in all_ch:
        return

    ins = all_ch.index(before_elem)
    offset = 0
    for _, crs, _, _ in comments:
        if crs is not None:
            p_elem.insert(ins + offset, crs)
            offset += 1

    all_ch2 = list(p_elem)
    if after_elem is None or after_elem not in all_ch2:
        return
    after_idx = all_ch2.index(after_elem) + 1
    for _, _, cre, crr in comments:
        if cre is not None:
            p_elem.insert(after_idx, cre)
            after_idx += 1
        if crr is not None:
            p_elem.insert(after_idx, crr)
            after_idx += 1


def _apply_replacements(para, replacements: list):
    """
    replacements: list of (block_text, field_instr_or_None, do_red, leftover_suffix)
    Processed in reverse order so earlier char positions stay valid.
    Non-run elements (comments, bookmarks) are never touched.
    leftover_suffix: text to insert as plain run immediately after the replacement
    (e.g. ' (figure 2.2.1)' for blocks like '(Francey et al 2003, figure 2.2.1)').
    """
    p_elem = para._p

    def build_map():
        runs = _get_plain_runs(para)
        text = ""
        cmap = []
        for r in runs:
            t = _get_run_text(r)
            for j in range(len(t)):
                cmap.append((r, j))
            text += t
        return text, cmap

    for entry in reversed(replacements):
        # Unpack: 5-tuple (linked), 6-tuple (old mixed), 8-tuple (red or new mixed)
        reasons = None
        if len(entry) == 8:
            block, field_instr, do_red, leftover, _clean, red_text, _, reasons = entry
            is_mixed = (field_instr is not None)
        elif len(entry) == 6:
            block, field_instr, do_red, leftover, _clean, red_text = entry
            is_mixed = True
        else:
            block, field_instr, do_red, leftover, _clean = entry
            red_text = _clean
            is_mixed = False

        full_text, char_map = build_map()
        pos = full_text.rfind(block)
        if pos == -1:
            continue
        end = pos + len(block) - 1  # inclusive

        start_run, start_off = char_map[pos]
        end_run,   end_off   = char_map[end]

        same_run = (start_run is end_run)

        if same_run:
            run_text  = _get_run_text(start_run)
            after_off = end_off + 1
            if after_off < len(run_text):
                start_run, _ = _split_run_at(start_run, after_off)
            if start_off > 0:
                _, start_run = _split_run_at(start_run, start_off)
            end_run = start_run
        else:
            end_run_text = _get_run_text(end_run)
            if end_off + 1 < len(end_run_text):
                end_run, _ = _split_run_at(end_run, end_off + 1)
            if start_off > 0:
                _, start_run = _split_run_at(start_run, start_off)

        # Refresh after splits
        all_children = list(p_elem)
        i_start = all_children.index(start_run)
        i_end   = all_children.index(end_run)

        # Exclude runs that carry comment reference marks — these must never be removed
        runs_to_replace = [
            c for c in all_children[i_start:i_end + 1]
            if c.tag == qn("w:r")
            and c.find(qn("w:commentReference")) is None
        ]

        if not runs_to_replace:
            continue

        ref_rpr   = runs_to_replace[0].find(qn("w:rPr"))
        insert_at = list(p_elem).index(runs_to_replace[0])

        # Collect comment elements overlapping this span; remove from current positions
        span_comments = _comment_ids_on_span(p_elem, i_start, i_end)
        for _, crs, cre, crr in span_comments:
            for e in (crs, cre, crr):
                if e is not None and e in list(p_elem):
                    p_elem.remove(e)

        # Build new content elements
        if is_mixed:
            new_elems = _make_field_elements(field_instr, _clean, ref_rpr)
            if reasons:
                new_elems.append(_make_plain_run("(", ref_rpr))
                for idx_r, (matched, tag, _) in enumerate(reasons):
                    rr = _make_plain_run(f"{matched} {tag}", ref_rpr)
                    _set_red_text(rr)
                    new_elems.append(rr)
                    if idx_r < len(reasons) - 1:
                        new_elems.append(_make_plain_run("; ", ref_rpr))
                new_elems.append(_make_plain_run(")", ref_rpr))
            else:
                red_run = _make_plain_run(red_text, ref_rpr)
                _set_red_text(red_run)
                new_elems.append(red_run)
            if leftover:
                new_elems.append(_make_plain_run(f" ({leftover})", ref_rpr))
        elif do_red:
            if reasons:
                new_elems = [_make_plain_run("(", ref_rpr)]
                for idx_r, (matched, tag, _) in enumerate(reasons):
                    rr = _make_plain_run(f"{matched} {tag}", ref_rpr)
                    _set_red_text(rr)
                    new_elems.append(rr)
                    if idx_r < len(reasons) - 1:
                        new_elems.append(_make_plain_run("; ", ref_rpr))
                new_elems.append(_make_plain_run(")", ref_rpr))
                runs_to_replace = runs_to_replace  # all get removed
            else:
                _set_run_text(runs_to_replace[0], red_text)
                _set_red_text(runs_to_replace[0])
                new_elems = [runs_to_replace[0]]
                runs_to_replace = runs_to_replace[1:]
            if leftover:
                new_elems.append(_make_plain_run(f" ({leftover})", ref_rpr))
        else:
            new_elems = _make_field_elements(field_instr, _clean, ref_rpr)
            if leftover:
                new_elems.append(_make_plain_run(f" ({leftover})", ref_rpr))

        # Remove covered plain runs
        if do_red and not reasons:
            for r in runs_to_replace:  # extras after [0]
                p_elem.remove(r)
        else:
            for r in runs_to_replace:
                p_elem.remove(r)

        # Insert new elements
        insert_pos = insert_at
        for i, elem in enumerate(new_elems):
            p_elem.insert(insert_pos + i, elem)

        # Reposition comment elements around the appropriate content
        if not span_comments:
            pass
        elif is_mixed and reasons:
            # Mixed case: attribute each comment to linked or red part by overlap
            all_ch = list(p_elem)
            # Find first/last field element (linked) and first/last red run
            field_elems = [e for e in new_elems
                           if e.tag == qn("w:r") and e in all_ch
                           and e.find(qn("w:fldChar")) is not None]
            red_run_elems = [
                e for e in new_elems
                if e.tag == qn("w:r") and e in all_ch
                and e.find(qn("w:t")) is not None
                and e.find(qn("w:t")).text is not None
                and "REF:" in e.find(qn("w:t")).text
            ]
            # Get index spans of linked and red parts
            if field_elems and red_run_elems:
                linked_span_now = (all_ch.index(field_elems[0]),
                                   all_ch.index(field_elems[-1]))
                red_span_now    = (all_ch.index(red_run_elems[0]),
                                   all_ch.index(red_run_elems[-1]))
                linked_coms, red_coms = _attribute_comments_to_spans(
                    p_elem, span_comments, linked_span_now, red_span_now)
                _reinsert_comments(p_elem, linked_coms,
                                   field_elems[0], field_elems[-1])
                # Reinsert red comments per-run by proximity
                for cid, crs, cre, crr in red_coms:
                    all_ch2 = list(p_elem)
                    cs_idx  = all_ch2.index(crs) if crs is not None and crs in all_ch2 else -1
                    ce_idx  = all_ch2.index(cre) if cre is not None and cre in all_ch2 else cs_idx
                    mid     = (cs_idx + ce_idx) / 2 if cs_idx >= 0 else 0
                    best_run  = None
                    best_dist = float("inf")
                    for rr in red_run_elems:
                        if rr not in all_ch2:
                            continue
                        ri = all_ch2.index(rr)
                        d  = abs(ri - mid)
                        if d < best_dist:
                            best_dist = d
                            best_run  = rr
                    if best_run is not None:
                        _reinsert_comments(p_elem, [(cid, crs, cre, crr)],
                                           best_run, best_run)
            else:
                # Fallback: wrap everything
                first_el = new_elems[0] if new_elems else None
                last_el  = new_elems[-1] if new_elems else None
                _reinsert_comments(p_elem, span_comments, first_el, last_el)
        elif reasons:
            # Pure red with per-ref runs: attribute each comment to its specific red run
            all_ch = list(p_elem)
            # Build matched_text → red run element map
            red_run_map: dict[str, object] = {}
            for matched, tag, _ in reasons:
                run_text = f"{matched} {tag}"
                for e in new_elems:
                    if (e.tag == qn("w:r") and e in all_ch
                            and e.find(qn("w:t")) is not None
                            and e.find(qn("w:t")).text == run_text):
                        red_run_map[matched] = e
                        break

            if red_run_map:
                # For each comment, find which red run its matched_text belongs to
                # by checking overlap with the original span of that ref.
                # Since we lost per-ref spans, we use the comment's position
                # relative to the red runs' positions to attribute it.
                # Strategy: attribute to the red run whose child index is closest
                # to the midpoint of the comment's [cs, ce] span.
                for cid, crs, cre, crr in span_comments:
                    all_ch2  = list(p_elem)
                    cs_idx   = all_ch2.index(crs) if crs is not None and crs in all_ch2 else -1
                    ce_idx   = all_ch2.index(cre) if cre is not None and cre in all_ch2 else cs_idx
                    mid      = (cs_idx + ce_idx) / 2 if cs_idx >= 0 else 0
                    # Find closest red run by index
                    best_run = None
                    best_dist = float("inf")
                    for matched, rr in red_run_map.items():
                        if rr not in all_ch2:
                            continue
                        ri = all_ch2.index(rr)
                        d  = abs(ri - mid)
                        if d < best_dist:
                            best_dist = d
                            best_run  = rr
                    if best_run is not None:
                        _reinsert_comments(p_elem, [(cid, crs, cre, crr)],
                                           best_run, best_run)
            else:
                first_el = new_elems[0] if new_elems else None
                last_el  = new_elems[-1] if new_elems else None
                _reinsert_comments(p_elem, span_comments, first_el, last_el)
        else:
            # Single block (pure linked or single red): wrap all around new content
            first_el = new_elems[0] if new_elems else None
            last_el  = new_elems[-1] if new_elems else None
            _reinsert_comments(p_elem, span_comments, first_el, last_el)


def _ref_in_context(ref: dict, all_refs: list) -> str:
    """
    Build a display string for one ref within the full block context.
    Each unique matched_text is shown once; others become '...'.
    For same-matched_text refs (same-author multi-year like Li 2024/2025),
    the specific year is highlighted within the shared matched_text.
    """
    target_mt = ref.get("matched_text", f"{ref['author']} {ref['year']}")
    same_mt_refs = [r for r in all_refs if r.get("matched_text") == target_mt]
    multi_year   = len(same_mt_refs) > 1

    parts = []
    seen_matched = set()
    for r in all_refs:
        mt = r.get("matched_text", f"{r['author']} {r['year']}")
        if mt in seen_matched:
            continue
        seen_matched.add(mt)
        if mt == target_mt:
            if multi_year:
                other_years = [x.get("year", "") for x in same_mt_refs if x is not ref]
                display_mt = mt
                for oy in other_years:
                    display_mt = re.sub(r',?\s*' + re.escape(oy) + r'[a-z]?', ',…', display_mt)
                display_mt = re.sub(r',…,', ',…', display_mt)
                parts.append(display_mt.strip().strip(',').strip())
            else:
                parts.append(mt)
        else:
            parts.append("...")
    deduped = []
    for p in parts:
        if p == "..." and deduped and deduped[-1] == "...":
            continue
        deduped.append(p)
    return "(" + "; ".join(deduped) + ")"


def _zotero_label(item: dict) -> str:
    """Format 'LastName1 and LastName2, YEAR' or 'LastName et al., YEAR'
    using the actual names and date stored in Zotero."""
    creators = item.get("creators", [])
    fields   = item.get("fields", {})
    year_str = fields.get("date", fields.get("year", ""))[:4]
    if len(creators) == 0:
        author_str = "?"
    elif len(creators) == 1:
        author_str = creators[0].get("lastName", "?")
    elif len(creators) == 2:
        author_str = (f"{creators[0].get('lastName','?')}"
                      f" and {creators[1].get('lastName','?')}")
    else:
        author_str = creators[0].get("lastName", "?") + " et al."
    return f"{author_str}, {year_str}"


# ─── Interactive choice ───────────────────────────────────────────────────

def _clear_lines(n: int):
    """Erase the last n lines in the terminal using ANSI codes."""
    for _ in range(n):
        sys.stdout.write("\x1b[1A\x1b[2K")
    sys.stdout.flush()


def _sentence_context(full_text: str, block: str, ref: dict) -> str:
    """
    Return the sentence containing block, truncated to the shorter of:
      (1) the end of the sentence
      (2) 60 characters after the end of the citation block
    The specific ref is highlighted with ** markers.
    """
    pos = full_text.find(block)
    if pos == -1:
        return full_text[:160]

    # Sentence start: last '.' before the block
    start = max(0, full_text.rfind('.', 0, pos) + 1)

    # Sentence end option (1): next '.' after block
    end_sent = full_text.find('.', pos + len(block))
    end_sent = end_sent + 1 if end_sent != -1 else len(full_text)

    # Sentence end option (2): 60 chars after end of citation
    end_60 = pos + len(block) + 60

    end = min(end_sent, end_60)
    sentence = full_text[start:end].strip()
    if end < end_sent:
        sentence += "…"

    # Build the highlighted version of the citation block
    matched_txt = ref.get("matched_text", "")
    year        = ref["year"] + ref.get("year_suffix", "")

    if matched_txt and matched_txt in block:
        # Find all years in matched_txt
        all_years = re.findall(_YEAR_PAT, matched_txt)
        if len(all_years) > 1:
            # Same-author multi-year: bold the author name + specific year only
            first_year_match = re.search(_YEAR_PAT, matched_txt)
            author_part_raw  = matched_txt[:first_year_match.start()]  # "Li et al., "
            author_part_trim = author_part_raw.rstrip(", ")            # "Li et al."
            sep_after_author = author_part_raw[len(author_part_trim):]  # ", "
            years_part       = matched_txt[first_year_match.start():]   # "2024,2025"
            # Bold only the specific year (its leading comma stays outside **)
            year_in_mt = re.search(r'(,?\s*)(' + re.escape(year) + r')', years_part)
            if year_in_mt:
                sep  = year_in_mt.group(1)   # leading comma/space before this year
                # The years before this one (between author and this year)
                years_before = years_part[:year_in_mt.start()]   # e.g. "2024," for Li 2025
                years_after  = matched_txt[len(author_part_raw) + year_in_mt.end():]  # remaining
                # Bold: just author + sep_after_author + years_before (plain) is wrong
                # Correct: bold only [author_part_trim + sep_after_author + years_before + sep + year]
                # But per user spec, for Li 2025: **Li et al.**, 2024,**2025**
                # So bold author separately, then bold just the year
                if years_before.strip(',').strip():
                    # There are years between author and the target year
                    # Show: **author_part_trim**,years_before,**year**
                    highlighted_block = block.replace(
                        matched_txt,
                        f"**{author_part_trim}**{sep_after_author}{years_before}{sep}**{year}**{years_after}"
                    )
                else:
                    # Target year is the first year → **author_part_trim, year**rest
                    highlighted_block = block.replace(
                        matched_txt,
                        f"**{author_part_trim}{sep_after_author}{sep}{year}**{years_after}"
                    )
            else:
                highlighted_block = block.replace(matched_txt, f"**{matched_txt}**")
        else:
            # Single year: bold the entire matched_text
            highlighted_block = block.replace(matched_txt, f"**{matched_txt}**")
    else:
        highlighted_block = block  # fallback: no highlighting

    # Replace the block in the sentence
    highlighted_sentence = sentence.replace(block, highlighted_block, 1)
    return highlighted_sentence


def interactive_choice(ref: dict, candidates: list[dict], full_text: str,
                       original_block: str, all_refs: list,
                       para_comments: list[str]) -> dict | None:
    author, year = ref["author"], ref["year"]
    context = _sentence_context(full_text, original_block, ref)

    lines_printed = 0

    print(f"\n{'─'*60}")
    lines_printed += 2
    print(f"  Context: \"{context}\"")
    lines_printed += 1
    if para_comments:
        for c in para_comments:
            print(f"  💬 {c}")
            lines_printed += 1
    else:
        print(f"  💬 No comment detected")
        lines_printed += 1
    print(f"  Homonym: {author} et al., {year}")
    lines_printed += 1
    print(f"  0) Skip (red)")
    lines_printed += 1
    for i, item in enumerate(candidates, 1):
        title       = item["fields"].get("title", "(no title)")
        journal     = item["fields"].get("publicationTitle", "")
        date        = item["fields"].get("date", "")[:4]
        authors_str = ", ".join(c["lastName"] for c in item["creators"][:3])
        if len(item["creators"]) > 3:
            authors_str += " et al."
        line = f"  {i}) {authors_str} ({date}) — {title[:55]}"
        if journal:
            line += f"  [{journal[:30]}]"
        print(line)
        lines_printed += 1

    choice_prompt = f"  Choice [0-{len(candidates)}]: "
    lines_printed += 1

    while True:
        try:
            raw = input(choice_prompt).strip()
            n   = int(raw)
            if 0 <= n <= len(candidates):
                break
        except (ValueError, KeyboardInterrupt):
            pass
        _clear_lines(1)

    _clear_lines(lines_printed)

    if n == 0:
        ref_display = _ref_in_context(ref, all_refs)
        print(f"  HOMO SKIPPED  {ref_display}")
        return None
    else:
        chosen      = candidates[n - 1]
        ref_display = _ref_in_context(ref, all_refs)
        zot_label   = _zotero_label(chosen)
        title       = chosen["fields"].get("title", "")[:40]
        print(f"  HOMO LINKED   {ref_display}  →  {zot_label}, {title}")
        return chosen


def _format_candidates(ref_label: str, candidates: list[dict]) -> str:
    """
    Format the list of Zotero homonym candidates as a comment text,
    using the same style as the interactive choice menu.
    """
    lines = [f"Multiple matches for: {ref_label}"]
    for i, item in enumerate(candidates, 1):
        title       = item["fields"].get("title", "(no title)")
        journal     = item["fields"].get("publicationTitle", "")
        date        = item["fields"].get("date", "")[:4]
        authors_str = ", ".join(c["lastName"] for c in item["creators"][:3])
        if len(item["creators"]) > 3:
            authors_str += " et al."
        line = f"{i}) {authors_str} ({date}) — {title[:55]}"
        if journal:
            line += f"  [{journal[:30]}]"
        lines.append(line)
    return "\n".join(lines)


def _add_word_comment(para, anchor_text: str, comment_text: str,
                      comments_map: dict) -> None:
    """
    Add a Word comment anchored to the run(s) containing anchor_text in para.
    Modifies the paragraph XML in-place and updates comments_map with the new entry.
    The comment is written into the document's comments part.
    """
    from lxml import etree as ET

    # Find the document part to access comments
    p_elem  = para._p
    # Walk up to find the document element
    doc_el  = p_elem.getroottree().getroot()

    # Get or create the comments part via the relationship
    # python-docx doesn't expose comments directly; we work with the XML tree
    # through the part stored on the paragraph's parent document.
    # We access it via the _element's ownerDocument trick.
    try:
        body = p_elem.getparent()
        document_el = body.getparent()
        # The document part is accessible via the namespace map
    except Exception:
        return

    # Find the comments part in the package
    # We need to get to the Document object — find it via the para's part
    # python-docx stores the part reference accessible through the element
    # Use a simpler approach: find existing comment XML or create it
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Determine a new unique comment ID
    existing_ids = [int(cid) for cid in comments_map.keys() if cid.isdigit()]
    new_id = str(max(existing_ids, default=-1) + 1)

    # Find the run containing anchor_text
    target_run = None
    for r in _get_plain_runs(para):
        if anchor_text in _get_run_text(r):
            target_run = r
            break
    if target_run is None:
        return

    # Build comment XML elements in the paragraph
    # 1. commentRangeStart — insert before target_run
    crs = OxmlElement("w:commentRangeStart")
    crs.set(qn("w:id"), new_id)
    p_elem.insert(list(p_elem).index(target_run), crs)

    # 2. commentRangeEnd — insert after target_run
    idx_after = list(p_elem).index(target_run) + 1
    cre = OxmlElement("w:commentRangeEnd")
    cre.set(qn("w:id"), new_id)
    p_elem.insert(idx_after, cre)

    # 3. commentReference run — insert after commentRangeEnd
    cr_run = OxmlElement("w:r")
    cr_rpr = OxmlElement("w:rPr")
    cr_style = OxmlElement("w:rStyle")
    cr_style.set(qn("w:val"), "CommentReference")
    cr_rpr.append(cr_style)
    cr_run.append(cr_rpr)
    cr_ref = OxmlElement("w:commentReference")
    cr_ref.set(qn("w:id"), new_id)
    cr_run.append(cr_ref)
    p_elem.insert(idx_after + 1, cr_run)

    # 4. Build the <w:comment> element and add it to comments_map for later saving
    comment_el = ET.Element(f"{{{W}}}comment")
    comment_el.set(f"{{{W}}}id",     new_id)
    comment_el.set(f"{{{W}}}author", "Zotero Linker")
    comment_el.set(f"{{{W}}}date",   "2024-01-01T00:00:00Z")

    for line in comment_text.split("\n"):
        cp = ET.SubElement(comment_el, f"{{{W}}}p")
        cr = ET.SubElement(cp, f"{{{W}}}r")
        ct = ET.SubElement(cr, f"{{{W}}}t")
        ct.text = line
        ct.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Store the element and text in comments_map using a special key format
    comments_map[new_id] = comment_text
    # Store the XML element for later writing — use a separate dict on comments_map
    if "_pending_elements" not in comments_map:
        comments_map["_pending_elements"] = {}
    comments_map["_pending_elements"][new_id] = comment_el


# ─── Paragraph processing ─────────────────────────────────────────────────

def process_paragraph(para, db: ZoteroDB, stats: dict, resolved_cache: dict,
                      homo_mode: str, show_linked: bool, para_num: int,
                      comments_map: dict):
    # Count existing Zotero fields (don't skip — paragraph may also have manual citations)
    existing_fields = _count_zotero_fields(para)
    stats["already_linked"] += existing_fields

    plain_runs = _get_plain_runs(para)
    if not plain_runs:
        return
    full_text = "".join(_get_run_text(r) for r in plain_runs)

    blocks = find_citation_blocks(full_text)
    if not blocks:
        return

    # Paragraph header — less indented than citation messages
    preview = full_text.strip()[:30].replace("\n", " ")
    print(f'§{para_num}: "{preview}..."')

    # replacements: list of replacement tuples built per block
    replacements          = []
    # multi_choice_comments: list of (red_run_label, candidates) for post-run comment insertion
    multi_choice_comments = []

    for original_block, refs, clean_inner, leftover in blocks:
        clean_block = f"({clean_inner})" if clean_inner else original_block

        # Resolve every individual ref in this block
        resolved_items    = []   # (ref, item) for successfully found refs
        failed_refs       = []   # (ref, reason) for not_found / ignored
        all_refs_in_block = list(refs)  # kept for context display in interactive_choice

        for ref in refs:
            cache_key = f"{normalize(ref['author'])}|{ref['year']}"
            was_interactive = False

            if cache_key in resolved_cache and (homo_mode in ("memo", "never") or resolved_cache[cache_key] == "not_found"):
                decision = resolved_cache[cache_key]
            else:
                candidates = db.search(ref["author"], ref["year"], ref.get("author2", ""))
                if len(candidates) == 0:
                    decision = "not_found"
                    resolved_cache[cache_key] = decision
                elif len(candidates) == 1:
                    decision = candidates[0]
                    resolved_cache[cache_key] = decision
                else:
                    if homo_mode == "never":
                        # Auto-skip without asking
                        decision = "ignored"
                        resolved_cache[cache_key] = decision
                        print(f"  HOMO SKIPPED  ({ref['author']} {ref['year']})  [auto]")
                    else:
                        was_interactive = True
                        ref_span = ref.get("matched_text", original_block)
                        block_comments = _comments_on_block(para, full_text, ref_span, comments_map)
                        chosen = interactive_choice(ref, candidates, full_text, original_block, all_refs_in_block, block_comments)
                        decision = chosen if chosen is not None else "ignored"
                        if homo_mode == "memo":
                            resolved_cache[cache_key] = decision

            if decision == "not_found":
                failed_refs.append((ref, "not_found", []))
                stats["not_found"] += 1
                matched = ref.get("matched_text", f"{ref['author']} {ref['year']}")
                ref_comments = _comments_on_block(
                    para, full_text,
                    ref.get("matched_text", original_block),
                    comments_map
                )
                if matched not in stats["missing"]:
                    stats["missing"][matched] = []
                for c in ref_comments:
                    if c not in stats["missing"][matched]:
                        stats["missing"][matched].append(c)
                print(f"  NOT FOUND     ({matched})")
            elif decision == "ignored":
                # Store candidates so we can attach them as a comment later
                cands = candidates if 'candidates' in dir() and len(candidates) > 1 else []
                failed_refs.append((ref, "ignored", cands))
                stats["ignored"] += 1
                # HOMO SKIPPED already printed by interactive_choice
            else:
                resolved_items.append((ref, decision, was_interactive))

        def _clean_for_resolved(resolved: list) -> str:
            """Return clean display text containing only the resolved refs.
            Uses the 'matched_text' stored per ref, deduplicating same matched_text
            (which happens for same-author multi-year refs like Li 2024 and Li 2025
            both pointing to 'Li et al., 2024,2025').
            """
            seen  = set()
            parts = []
            for ref, item, _ in resolved:
                txt = ref.get("matched_text", f"{ref['author']} {ref['year']}")
                if txt not in seen:
                    seen.add(txt)
                    parts.append(txt)
            return "; ".join(parts)

        if resolved_items and not failed_refs:
            # All linked → single clean Zotero field
            items_only   = [item for _, item, _ in resolved_items]
            linked_clean = f"({_clean_for_resolved(resolved_items)})"
            field_instr  = build_zotero_field_xml(items_only, linked_clean)
            replacements.append((original_block, field_instr, False, leftover, linked_clean))
            stats["replaced"] += len(resolved_items)
            if show_linked:
                for ref, item, was_interactive in resolved_items:
                    if not was_interactive:
                        title       = item["fields"].get("title", "")[:40]
                        ref_display = _ref_in_context(ref, all_refs_in_block)
                        zot_label   = _zotero_label(item)
                        print(f"  LINKED        {ref_display}  →  {zot_label}, {title}")

        elif not resolved_items and failed_refs:
            # All failed → per-ref red runs
            reasons  = []
            for ref, reason, cands in failed_refs:
                tag     = "REF: NOT FOUND" if reason == "not_found" else "REF: MULTI CHOICE"
                matched = ref.get("matched_text", f"{ref['author']} {ref['year']}")
                reasons.append((matched, tag, cands))
            red_text = "(" + "; ".join(f"{m} {t}" for m, t, _ in reasons) + ")"
            # 8-tuple: include reasons for per-ref comment anchoring
            replacements.append((original_block, None, True, leftover, red_text, None, None, reasons))
            multi_choice_comments.extend(
                (f"{m} REF: MULTI CHOICE", c) for m, t, c in reasons
                if t == "REF: MULTI CHOICE" and c
            )

        else:
            # Mixed: Zotero field for linked refs only + per-ref red runs for failed refs
            items_only   = [item for _, item, _ in resolved_items]
            linked_clean = f"({_clean_for_resolved(resolved_items)})"
            field_instr  = build_zotero_field_xml(items_only, linked_clean)
            reasons = []
            for ref, reason, cands in failed_refs:
                tag     = "REF: NOT FOUND" if reason == "not_found" else "REF: MULTI CHOICE"
                matched = ref.get("matched_text", f"{ref['author']} {ref['year']}")
                reasons.append((matched, tag, cands))
            red_text = "(" + "; ".join(f"{m} {t}" for m, t, _ in reasons) + ")"
            # 8-tuple for mixed case
            replacements.append((original_block, field_instr, None, leftover,
                                  linked_clean, red_text, None, reasons))
            stats["replaced"] += len(resolved_items)
            multi_choice_comments.extend(
                (f"{m} REF: MULTI CHOICE", c) for m, t, c in reasons
                if t == "REF: MULTI CHOICE" and c
            )
            if show_linked:
                for ref, item, was_interactive in resolved_items:
                    if not was_interactive:
                        title       = item["fields"].get("title", "")[:40]
                        ref_display = _ref_in_context(ref, all_refs_in_block)
                        zot_label   = _zotero_label(item)
                        print(f"  LINKED        {ref_display}  →  {zot_label}, {title}")

    if replacements:
        _apply_replacements(para, replacements)

    # Add Word comments for MULTI CHOICE refs, anchored to the red run
    for red_run_label, candidates in multi_choice_comments:
        comment_text = _format_candidates(red_run_label, candidates)
        _add_word_comment(para, red_run_label, comment_text, comments_map)


# ─── Interactive startup prompts ──────────────────────────────────────────

def prompt(label: str, default: str | None = None) -> str:
    """Prompt for a value. Strips surrounding quotes to allow paths with spaces."""
    def clean(v: str) -> str:
        v = v.strip()
        # Remove surrounding quotes (single or double) that users may add
        # when paths contain spaces, e.g. "my file.docx" or 'my file.docx'
        if len(v) >= 2 and v[0] in ('"', "'") and v[-1] == v[0]:
            v = v[1:-1]
        return v

    if default:
        value = clean(input(f"  {label} [{default}]: "))
        return value if value else default
    while True:
        value = clean(input(f"  {label}: "))
        if value:
            return value
        print("   Required.")


def _load_comments(docx_path: str) -> dict[str, str]:
    """
    Extract comments from word/comments.xml inside the docx.
    Returns a dict mapping comment id (str) → comment text (str).
    """
    import zipfile
    comments: dict = {}
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            if 'word/comments.xml' not in z.namelist():
                return comments
            from lxml import etree as ET
            xml  = z.read('word/comments.xml')
            root = ET.fromstring(xml)
            W    = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            for comment in root.findall(f'{{{W}}}comment'):
                cid  = comment.get(f'{{{W}}}id', '')
                text = "".join(
                    t.text or ""
                    for t in comment.iter(f'{{{W}}}t')
                )
                if cid and text.strip():
                    comments[cid] = text.strip()
    except Exception:
        pass
    return comments


def _write_pending_comments(out_path: str, comments_map: dict) -> None:
    """
    Write any pending comment elements (added during processing) into the
    saved docx file's word/comments.xml.
    """
    pending = comments_map.get("_pending_elements", {})
    if not pending:
        return

    import zipfile, shutil, os, tempfile
    from lxml import etree as ET

    W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS  = {"w": W}
    tmp = out_path + ".tmp_comments"

    with zipfile.ZipFile(out_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "word/comments.xml":
                    root = ET.fromstring(data)
                    for cid, el in pending.items():
                        root.append(el)
                    data = ET.tostring(root, xml_declaration=True,
                                       encoding="UTF-8", standalone=True)

                elif item.filename == "word/_rels/document.xml.rels":
                    # Ensure relationship to comments.xml exists
                    rel_root = ET.fromstring(data)
                    rel_ns   = "http://schemas.openxmlformats.org/package/2006/relationships"
                    comment_type = ("http://schemas.openxmlformats.org/officeDocument"
                                    "/2006/relationships/comments")
                    existing_types = [r.get("Type") for r in rel_root.findall(
                        f"{{{rel_ns}}}Relationship")]
                    if comment_type not in existing_types:
                        # Generate a new relationship ID
                        existing_rids = [r.get("Id","") for r in rel_root.findall(
                            f"{{{rel_ns}}}Relationship")]
                        rnum = max((int(r[2:]) for r in existing_rids
                                    if r.startswith("rId") and r[2:].isdigit()),
                                   default=10) + 1
                        rel = ET.SubElement(rel_root, f"{{{rel_ns}}}Relationship")
                        rel.set("Id",     f"rId{rnum}")
                        rel.set("Type",   comment_type)
                        rel.set("Target", "comments.xml")
                        data = ET.tostring(rel_root, xml_declaration=True,
                                           encoding="UTF-8", standalone=True)

                elif item.filename == "[Content_Types].xml":
                    # Ensure comments part is declared
                    ct_root = ET.fromstring(data)
                    ct_ns   = "http://schemas.openxmlformats.org/package/2006/content-types"
                    existing = [o.get("PartName") for o in ct_root.findall(
                        f"{{{ct_ns}}}Override")]
                    if "/word/comments.xml" not in existing:
                        ov = ET.SubElement(ct_root, f"{{{ct_ns}}}Override")
                        ov.set("PartName", "/word/comments.xml")
                        ov.set("ContentType",
                               "application/vnd.openxmlformats-officedocument"
                               ".wordprocessingml.comments+xml")
                        data = ET.tostring(ct_root, xml_declaration=True,
                                           encoding="UTF-8", standalone=True)

                zout.writestr(item, data)

            # If comments.xml didn't exist yet, create it
            names = [i.filename for i in zin.infolist()]
            if "word/comments.xml" not in names:
                root = ET.Element(f"{{{W}}}comments")
                for cid, el in pending.items():
                    root.append(el)
                data = ET.tostring(root, xml_declaration=True,
                                   encoding="UTF-8", standalone=True)
                zout.writestr("word/comments.xml", data)

    shutil.move(tmp, out_path)
    """
    Extract comments from word/comments.xml inside the docx.
    Returns a dict mapping comment id (str) → comment text (str).
    """
    import zipfile
    comments = {}
    try:
        with zipfile.ZipFile(docx_path, 'r') as z:
            if 'word/comments.xml' not in z.namelist():
                return comments
            from lxml import etree as ET
            xml = z.read('word/comments.xml')
            root = ET.fromstring(xml)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for comment in root.findall('w:comment', ns):
                cid  = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', '')
                text = "".join(t.text or "" for t in comment.iter(
                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
                if cid and text.strip():
                    comments[cid] = text.strip()
    except Exception:
        pass
    return comments


def _comment_ids_on_span(p_elem, i_start: int, i_end: int) -> list[tuple]:
    """
    Return list of (cid, commentRangeStart_elem, commentRangeEnd_elem, commentRef_run)
    for comments whose range overlaps [i_start, i_end] in p_elem's children.
    Missing elements are None.
    """
    all_children = list(p_elem)
    start_elems: dict[str, object] = {}
    end_elems:   dict[str, object] = {}
    ref_runs:    dict[str, object] = {}
    start_idx:   dict[str, int]    = {}
    end_idx:     dict[str, int]    = {}

    for i, child in enumerate(all_children):
        if child.tag == qn("w:commentRangeStart"):
            cid = child.get(qn("w:id"), "")
            if cid:
                start_elems[cid] = child
                start_idx[cid]   = i
        elif child.tag == qn("w:commentRangeEnd"):
            cid = child.get(qn("w:id"), "")
            if cid:
                end_elems[cid] = child
                end_idx[cid]   = i
        elif child.tag == qn("w:r"):
            cr = child.find(qn("w:commentReference"))
            if cr is not None:
                cid = cr.get(qn("w:id"), "")
                if cid:
                    ref_runs[cid] = child

    result = []
    for cid, cs in start_idx.items():
        ce = end_idx.get(cid, cs)
        # Only process comments that overlap [i_start, i_end] but do NOT
        # fully contain it (those that fully contain are left untouched).
        overlaps  = ce >= i_start and cs <= i_end
        fully_contains = cs <= i_start and ce >= i_end
        if overlaps and not fully_contains:
            result.append((
                cid,
                start_elems.get(cid),
                end_elems.get(cid),
                ref_runs.get(cid),
            ))
    return result


def _comments_on_block(para, full_text: str, block: str,
                        comments_map: dict) -> list[str]:
    """
    Return comment texts whose range overlaps (partially or totally) the span
    of plain-text runs covering 'block'.

    A comment [commentRangeStart(id=X) ... commentRangeEnd(id=X)] overlaps
    the citation span [i_start, i_end] if and only if:
        start of comment <= i_end   AND   end of comment >= i_start
    i.e. they are not disjoint.
    """
    if not comments_map:
        return []

    p_elem = para._p
    plain_runs = _get_plain_runs(para)
    if not plain_runs:
        return []

    # Build char map to find which runs cover the block
    text = ""
    cmap = []
    for r in plain_runs:
        t = _get_run_text(r)
        for j in range(len(t)):
            cmap.append((r, j))
        text += t

    pos = text.find(block)
    if pos == -1:
        return []
    end_pos = pos + len(block) - 1

    start_run = cmap[pos][0]
    end_run   = cmap[end_pos][0]

    all_children = list(p_elem)
    try:
        i_start = all_children.index(start_run)
        i_end   = all_children.index(end_run)
    except ValueError:
        return []

    # Build a map of comment id → (range_start_idx, range_end_idx) in all_children
    comment_start_idx: dict[str, int] = {}
    comment_end_idx:   dict[str, int] = {}
    for i, child in enumerate(all_children):
        if child.tag == qn("w:commentRangeStart"):
            cid = child.get(qn("w:id"), "")
            if cid:
                comment_start_idx[cid] = i
        elif child.tag == qn("w:commentRangeEnd"):
            cid = child.get(qn("w:id"), "")
            if cid:
                comment_end_idx[cid] = i

    # A comment overlaps [i_start, i_end] iff its range intersects that interval
    result = []
    for cid, cs in comment_start_idx.items():
        ce = comment_end_idx.get(cid, cs)  # if no end found, treat as point
        # Overlap condition: not (ce < i_start or cs > i_end)
        if ce >= i_start and cs <= i_end:
            if cid in comments_map:
                result.append(comments_map[cid])
    return result


def ask_parameters() -> dict:
    print("=" * 60)
    print("  Zotero Citation Linker")
    print("=" * 60)
    print()

    # .docx source
    default_docx = "test_2-2.docx"
    while True:
        docx_raw  = prompt("Source .docx", default_docx)
        docx_path = Path(docx_raw).expanduser()
        if docx_path.exists():
            break
        print(f"   Not found: {docx_path}")

    # Zotero SQLite
    default_db = None
    for candidate in [
        Path.home() / "Zotero" / "zotero.sqlite",
        Path.home() / "snap" / "zotero-snap" / "common" / "Zotero" / "zotero.sqlite",
    ]:
        if candidate.exists():
            default_db = str(candidate)
            break

    while True:
        db_raw  = prompt("Zotero database (zotero.sqlite)", default_db)
        db_path = Path(db_raw).expanduser()
        if db_path.exists():
            break
        print(f"   Not found: {db_path}")

    # Output — loop until a non-existing path is given
    default_out = docx_path.stem + "_zotero" + docx_path.suffix
    while True:
        out_raw  = prompt("Output .docx", default_out)
        out_path = Path(out_raw).expanduser()
        if not out_path.exists():
            break
        print(f"   File already exists: {out_path}")
        print("   Please delete it or choose a different name.")
        default_out = out_raw  # keep the user's last input as the new default

    return {"docx_path": docx_path, "db_path": db_path, "out_path": out_path}


def pick_library(db_path: Path) -> int | None:
    libs = ZoteroDB.list_libraries(str(db_path))
    if not libs:
        print("No libraries found.")
        return None

    print("\nAvailable libraries:")
    for lib in libs:
        print(f"  [{lib['libraryID']}] {lib['name']}  ({lib['type']})")

    while True:
        try:
            choice = input("\n  Library ID (Enter = all): ").strip()
            if choice == "":
                return None
            lib_id = int(choice)
            if any(l["libraryID"] == lib_id for l in libs):
                name = next(l["name"] for l in libs if l["libraryID"] == lib_id)
                print(f"  → {name}\n")
                return lib_id
        except (ValueError, KeyboardInterrupt):
            pass
        print("  Invalid ID.")


def ask_memo_mode() -> str:
    """
    Ask how to handle homonyms (multiple Zotero matches for one citation).
    Returns one of:
      "never"  — never ask, always skip homonyms (mark red automatically)
      "memo"   — ask once per author+year, reuse the choice for duplicates
      "ask"    — ask every time a homonym is encountered
    """
    print("─" * 60)
    print("Homonym handling (when a citation matches multiple Zotero entries):")
    print("  1) Never link — always skip homonyms automatically (mark red)")
    print("  2) Ask once   — memorize choice, reuse for all identical occurrences")
    print("  3) Ask always — ask again at each occurrence")
    while True:
        try:
            c = input("  [1/2/3]: ").strip()
            if c == "1":
                print("  → Homonyms will be skipped automatically.\n")
                return "never"
            if c == "2":
                print("  → Choice will be memorized per author+year.\n")
                return "memo"
            if c == "3":
                print("  → You will be asked at each occurrence.\n")
                return "ask"
        except KeyboardInterrupt:
            sys.exit(0)
        print("  Enter 1, 2, or 3.")


def ask_show_linked() -> bool:
    print("─" * 60)
    print("Show each citation as it is processed?")
    print("  1) Yes — show linked citations and failures one by one")
    print("  2) No  — only show a summary at the end")
    while True:
        try:
            c = input("  [1/2]: ").strip()
            if c == "1":
                return True
            if c == "2":
                return False
        except KeyboardInterrupt:
            sys.exit(0)
        print("  Enter 1 or 2.")


# ─── Entry point ──────────────────────────────────────────────────────────

def main():
    try:
        params = ask_parameters()
    except KeyboardInterrupt:
        print("\nAborted.")
        sys.exit(0)

    docx_path = params["docx_path"]
    db_path   = params["db_path"]
    out_path  = params["out_path"]

    library_id  = pick_library(db_path)
    homo_mode   = ask_memo_mode()
    show_linked = ask_show_linked()

    print("─" * 60)

    db  = ZoteroDB(str(db_path), library_id)
    doc = Document(str(docx_path))

    # Load comments from the docx (word/comments.xml) for display during homonym choice
    comments_map = _load_comments(str(docx_path))

    stats = {"replaced": 0, "not_found": 0, "ignored": 0, "already_linked": 0,
             "missing": {}}
    resolved_cache = {}

    all_paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)

    print(f"\n{'═'*60}")
    print(f"  START")
    print(f"{'═'*60}")
    print(f"\nScanning {len(all_paragraphs)} paragraphs...\n")
    for para_num, para in enumerate(all_paragraphs, 1):
        process_paragraph(para, db, stats, resolved_cache, homo_mode, show_linked,
                          para_num, comments_map)

    db.close()
    doc.save(str(out_path))
    _write_pending_comments(str(out_path), comments_map)

    print(f"\n{'═'*60}")
    print(f"  END")
    print(f"{'═'*60}")
    print(f"\nSaved: {out_path}")
    print(f"  Linked        : {stats['replaced']}")
    print(f"  Already linked: {stats['already_linked']}")
    print(f"  Homonyms (red): {stats['ignored']}")
    print(f"  Not found (red): {stats['not_found']}")

    if stats["missing"]:
        try:
            show = input("\nShow missing references with their comments? [y/N]: ").strip().lower()
        except KeyboardInterrupt:
            show = "n"
        if show == "y":
            print("\n  Missing from Zotero:")
            for matched in sorted(stats["missing"].keys()):
                print(f"    • {matched}")
                for c in stats["missing"][matched]:
                    print(f"      - {c}")

    print(f"\nNext steps:")
    print(f"  * Open your Zotero app.")
    print(f"  * Open {out_path.name} in Word with the Zotero plugin installed,")
    print(f"    then click on \"Document Preferences\" and choose the right citation style,")
    print(f"    then click Refresh to finalize citations.")


if __name__ == "__main__":
    main()