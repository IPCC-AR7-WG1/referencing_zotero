# zotero_link_citations.py

Replaces "manually typed" citations in a `.docx` document with native Zotero fields, using your local Zotero SQLite database. 

You should close your input file before running the script otherwise you get an error message saying that ' Package not found at *.docx'. Your local Zotero app should be closed before the code finishes otherwise your linking might crush after spending hours working on it. 

After running the script, open your local Zotero App and the output file in Word with the Zotero plugin, verify your citation style and click **Add/Edit Bibliography**.

---

## Requirements

- Python 3.10+
- Zotero installed locally (the script reads `zotero.sqlite` directly — **Zotero must be closed** when running the script)
- The Zotero Word plugin installed in Microsoft Word

Install Python dependencies:

```bash
pip install python-docx lxml
```

---

## Usage

```bash
python zotero_link_citations.py
```

The script is fully interactive: it will ask for all parameters at startup.

---

## Startup prompts

| Prompt | Description |
|---|---|
| **Source.docx** | Your input Word document with the path if you are not running from the folder where the document is in. Paths with spaces are supported. |
| **Zotero database** | Path to `zotero.sqlite`. Auto-detected if found in `~/Zotero/`. |
| **Output.docx** | Path for the output file. Will not overwrite an existing file. |
| **Library** | Which Zotero library to search. The 'WGI AR7 General' library is selected by default. You need to choose your Chapter Library as the additional library. |
| **Homonym handling** | What to do when a citation matches multiple Zotero entries. |
| **Show citations** | Whether to print each linked citation in the terminal as it is processed. It seems that it shows the quotes in all cases. |

---

## Citation formats recognised

Citations must be inside parentheses. Supported forms include:

```
(Bereiter et al., 2015)
(Simmons et al., 2017; Gillett et al., 2021)
(Chan et al., 2019, 2024)
(Li and Paul, 2026)
(S. Szopa et al., 2026)
(von Schuckmann et al., 2020)
```

Non-citation content mixed into a citation block (e.g. `(as in TEXT-REF)` `(TEXT-REF figure 2.2.1)`, `(TEXT-REF, Miocene)`) is extracted and placed in a separate parenthesis immediately after the citation.

---

## Homonym handling

When a citation matches multiple Zotero entries (eg several possibilities for `(Wu et al. 2025 )`), You can choose to process them in 3 ways:

| Mode | Behaviour |
|---|---|
| **Never link** | Always skip automatically and mark red in output .docx |
| **Ask once** | Show a choice menu when a new multi choice appears; reuse the answer for duplicate occurrences. Not recommended. |
| **Ask always** | Show a choice menu at every occurrence of multi choice |

The choice menu shows the sentence containing the citation (with the reference highlighted), any Word comment anchored to it, and the list of Zotero candidates (with all authors, title, ...). Choosing **0** skips the citation and marks it red and add a comment in output .docx containing the list of Zotero candidates.

---

## Output document

**Linked citations** are replaced by native Zotero fields. Multiple citations in one parenthesis become a single multi-citation field.

**Unlinked citations** appear in red:
```
(Foster and Rahmstorf, 2025 REF: NOT FOUND)
(Chan et al., 2025 REF: MULTI CHOICE)
```
For `REF: MULTI CHOICE`, a Word comment listing all candidate Zotero entries is automatically added so you can resolve it manually later.

**Word comments** are preserved. Comments that overlap a citation are repositioned around the new content.

---

## After the script

Open the output file in Word with the Zotero plugin, then:

1. Click **Zotero → Document Preferences** and select your citation style.
2. Click **Add/Edit Bibliography** to generate a reference list or **Refresh** to reformat all citations and the existing reference list.

Red citations (`REF: NOT FOUND` or `REF: MULTI CHOICE`) must be resolved manually: either add the missing reference to your Zotero library and re-run the script, or replace the red text with a Zotero citation directly in Word.
