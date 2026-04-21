#!/usr/bin/env python3
"""
format_bc_manuals.py

Copies BC manual Google Docs into the BC_Manuals_Updated folder and applies:
  - Styled note box (background #EBE8E2, left border #8CB7C7 3pt) to every
    paragraph whose text starts with Note:, Notes:, or NOTE: (case-insensitive)
  - Footer: "Brazosport College | CIE | D2L New Content Experience Page"

Setup (one-time):
  1. Go to https://console.cloud.google.com and create (or select) a project.
  2. Enable the Google Docs API and Google Drive API for that project.
  3. Under APIs & Services > Credentials, create an OAuth 2.0 Client ID
     (Application type: Desktop app) and download it as credentials.json.
  4. Place credentials.json in the same directory as this script.
  5. Install dependencies:  pip install -r requirements.txt
  6. Run:  python format_bc_manuals.py
     A browser window will open once for authorization; after that a
     token.pickle file is saved so you won't be prompted again.
"""

import os
import pickle
import re

from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

# Source Google Doc IDs  →  display titles
DOCS = {
    "1sU8l8p6q4sFIjIHjuyIt29xAp-QzzVWzegmWmPtjpko": "Attendance Tool",
    "1Dfd2kzXtR4ACeygHKgttzmz98kIBE4LsdHK9EbMzBDE": "bestpractice_discussion_rubric_grading",
    "1VOKZO4eSTuo44D206ZAjXPNaluohwDw2Gw0njvxZHtE": "Create A Quiz",
    "1bOv8Um7zB2EU53Y_E7V8iBhLmMlWur_bK-r2pkW1Q74": "Creating a Question Pool for Quizzes",
    "1bUyx4ikKSJCWfTO7E6Opt9nkNDVIg5WEB4-700hGkIY": "CreatingZoomMeeting",
    "1kms3Tj5TfP_ayngBUVHGIsQ9e5DuVM_ch_cz1cqZxcs": "Creating Questions in the Question Library",
    "1NGgD2FxebXjSNiWMGcJElaw2hdOUQJqvkjWmFRUHeZk": "Creating Quizzes",
    "1nfFlkaoRDkpAHMEy5kGaYVP-JmQivvZPQUr1IH0DeTY": "AddingVideoNote",
}

OUTPUT_FOLDER_ID = "1NYpaEBB4UtUdA2ObHEcSpZwrLe_3jRPj"  # BC_Manuals_Updated
FOOTER_TEXT = "Brazosport College | CIE | D2L New Content Experience Page"

# Note-box colors (hex → 0-1 float)
_NOTE_BG = {"red": 0xEB / 255, "green": 0xE8 / 255, "blue": 0xE2 / 255}
_NOTE_BORDER = {"red": 0x8C / 255, "green": 0xB7 / 255, "blue": 0xC7 / 255}


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def get_credentials():
    creds = None
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as fh:
            creds = pickle.load(fh)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.pickle", "wb") as fh:
            pickle.dump(creds, fh)
    return creds


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _paragraph_text(paragraph):
    """Return the plain-text content of a paragraph element."""
    return "".join(
        elem.get("textRun", {}).get("content", "")
        for elem in paragraph.get("elements", [])
    )


def _is_note_paragraph(text):
    """True if the paragraph starts with Note:, Notes:, or NOTE: etc."""
    return bool(re.match(r"^\s*notes?:", text, re.IGNORECASE))


def _find_note_paragraphs(doc):
    """Return [(startIndex, endIndex), ...] for every note paragraph in body."""
    results = []
    for elem in doc.get("body", {}).get("content", []):
        if "paragraph" not in elem:
            continue
        if _is_note_paragraph(_paragraph_text(elem["paragraph"])):
            results.append((elem["startIndex"], elem["endIndex"]))
    return results


# ---------------------------------------------------------------------------
# Formatting requests
# ---------------------------------------------------------------------------

def _note_style_requests(start, end):
    """Build a batchUpdate request that styles one note paragraph."""
    return [
        {
            "updateParagraphStyle": {
                "range": {"startIndex": start, "endIndex": end},
                "paragraphStyle": {
                    "shading": {
                        "backgroundColor": {"color": {"rgbColor": _NOTE_BG}}
                    },
                    "borderLeft": {
                        "color": {"color": {"rgbColor": _NOTE_BORDER}},
                        "dashStyle": "SOLID",
                        "padding": {"magnitude": 6, "unit": "PT"},
                        "width": {"magnitude": 3, "unit": "PT"},
                    },
                    "indentStart": {"magnitude": 12, "unit": "PT"},
                    "spaceAbove": {"magnitude": 4, "unit": "PT"},
                    "spaceBelow": {"magnitude": 4, "unit": "PT"},
                },
                "fields": "shading,borderLeft,indentStart,spaceAbove,spaceBelow",
            }
        }
    ]


def apply_note_formatting(docs_svc, doc_id):
    """Style every note paragraph in doc_id. Returns the count styled."""
    doc = docs_svc.documents().get(documentId=doc_id).execute()
    note_paras = _find_note_paragraphs(doc)
    if not note_paras:
        return 0
    requests = []
    # Reverse order so earlier indices stay valid after each update
    for start, end in reversed(note_paras):
        requests.extend(_note_style_requests(start, end))
    docs_svc.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()
    return len(note_paras)


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------

def set_footer(docs_svc, doc_id):
    """Replace (or create) the default footer with the Brazosport College line."""
    doc = docs_svc.documents().get(documentId=doc_id).execute()
    doc_style = doc.get("documentStyle", {})
    requests = []

    if "defaultFooterId" not in doc_style:
        result = docs_svc.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [{"createFooter": {"type": "DEFAULT"}}]},
        ).execute()
        footer_id = result["replies"][0]["createFooter"]["footerId"]
    else:
        footer_id = doc_style["defaultFooterId"]
        # Clear existing footer content (delete from last paragraph to first)
        footer_content = (
            doc.get("footers", {}).get(footer_id, {}).get("content", [])
        )
        deletions = []
        for elem in footer_content:
            if "paragraph" not in elem:
                continue
            start, end = elem["startIndex"], elem["endIndex"]
            if end - start > 1:  # skip paragraphs that are just \n
                deletions.append((start, end - 1))
        for start, end in reversed(deletions):
            requests.append(
                {
                    "deleteContentRange": {
                        "range": {
                            "startIndex": start,
                            "endIndex": end,
                            "segmentId": footer_id,
                        }
                    }
                }
            )

    # Insert footer text at position 1 (after any structural character)
    requests.append(
        {
            "insertText": {
                "location": {"index": 1, "segmentId": footer_id},
                "text": FOOTER_TEXT,
            }
        }
    )
    # Center-align the footer paragraph
    requests.append(
        {
            "updateParagraphStyle": {
                "range": {
                    "startIndex": 1,
                    "endIndex": 1 + len(FOOTER_TEXT),
                    "segmentId": footer_id,
                },
                "paragraphStyle": {"alignment": "CENTER"},
                "fields": "alignment",
            }
        }
    )

    docs_svc.documents().batchUpdate(
        documentId=doc_id, body={"requests": requests}
    ).execute()


# ---------------------------------------------------------------------------
# Drive helpers
# ---------------------------------------------------------------------------

def copy_to_folder(drive_svc, src_id, title, folder_id):
    """Copy src_id into folder_id and return the new file's ID."""
    result = drive_svc.files().copy(
        fileId=src_id,
        body={"name": title, "parents": [folder_id]},
    ).execute()
    return result["id"]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    creds = get_credentials()
    docs_svc = build("docs", "v1", credentials=creds)
    drive_svc = build("drive", "v3", credentials=creds)

    print(f"Processing {len(DOCS)} document(s)...\n")
    for src_id, title in DOCS.items():
        print(f"  {title}")
        new_id = copy_to_folder(drive_svc, src_id, title, OUTPUT_FOLDER_ID)
        note_count = apply_note_formatting(docs_svc, new_id)
        set_footer(docs_svc, new_id)
        status = f"{note_count} note(s) styled" if note_count else "no notes found"
        print(f"    → copied  |  {status}  |  footer set")

    print("\nAll done! Check the BC_Manuals_Updated folder in Google Drive.")


if __name__ == "__main__":
    main()
