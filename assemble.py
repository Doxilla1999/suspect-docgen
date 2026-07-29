"""
Combines the generated .docx templates (in build/) with app_template.html
into a single self-contained suspect-docgen.html file.

Run generate_templates.py first, then this script:
    python3 generate_templates.py
    python3 assemble.py
"""
import base64
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BUILD_DIR = os.path.join(SCRIPT_DIR, "build")

TEMPLATE_FILES = {
    "__B64_DRUG115__": "drug_test_115_template.docx",
    "__B64_URINE__": "urine_referral_template.docx",
    "__B64_COURT__": "court_referral_template.docx",
    "__B64_M80__": "phone_consent_m80_template.docx",
    "__B64_PROFILE_PHOTOS__": "profile_and_photos_template.docx",
}


def b64_of(filename):
    path = os.path.join(BUILD_DIR, filename)
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def main():
    html_path = os.path.join(SCRIPT_DIR, "app_template.html")
    with open(html_path, encoding="utf-8") as f:
        html = f.read()

    for token, filename in TEMPLATE_FILES.items():
        html = html.replace(token, b64_of(filename))

    if "__B64_" in html:
        raise SystemExit("Leftover __B64_ token found — a template failed to embed. Aborting.")

    out_path = os.path.join(SCRIPT_DIR, "suspect-docgen.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"assembled {out_path} ({len(html):,} chars)")


if __name__ == "__main__":
    main()
