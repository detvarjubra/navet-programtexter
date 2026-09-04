"""Export program texts from the Expo GraphQL API to Markdown files.

Fetches all `programs` (with their `programOffers`) from Expo and writes
one Markdown file per program into `program/`. Existing files are never
overwritten so editorial work is protected.

Run with:
    python export_program_texts.py
"""

import os
import re
import sys

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

API_URL = "https://booking.navet.com/api/v3/graphql"
API_TOKEN = os.environ.get("API_TOKEN", "")

if not API_TOKEN:
    raise RuntimeError(
        "API_TOKEN saknas. Lägg till den i .env eller sätt miljövariabeln."
    )

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "program")

GET_PROGRAMS_QUERY = """
query GetPrograms($after: String) {
  programs(first: 50, after: $after) {
    nodes {
      id
      name
      color

      programOffers(first: 50) {
        nodes {
          offer {
            id
            name
            slug
            description
          }
        }
      }
    }

    pageInfo {
      hasNextPage
      endCursor
    }
  }
}
"""

# Characters that need explicit transliteration before punctuation is stripped.
CHAR_REPLACEMENTS = {
    "å": "a",
    "ä": "a",
    "ö": "o",
    "Å": "a",
    "Ä": "a",
    "Ö": "o",
}


def slugify(text):
    """Convert text to a filesystem-safe, deterministic slug."""
    for char, replacement in CHAR_REPLACEMENTS.items():
        text = text.replace(char, replacement)
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s-]", "", text)
    text = re.sub(r"[\s_]+", "-", text).strip("-")
    text = re.sub(r"-+", "-", text)
    return text


def fetch_programs():
    """Fetch all programs from the Expo GraphQL API, following pagination."""
    programs = []
    after = None
    headers = {
        "Authorization": f"Bearer {API_TOKEN}",
        "Content-Type": "application/json",
    }

    while True:
        response = requests.post(
            API_URL,
            headers=headers,
            json={"query": GET_PROGRAMS_QUERY, "variables": {"after": after}},
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()

        if "errors" in payload and payload["errors"]:
            raise RuntimeError(f"GraphQL-fel: {payload['errors']}")

        page = payload["data"]["programs"]
        programs.extend(page["nodes"])

        page_info = page["pageInfo"]
        if page_info["hasNextPage"]:
            after = page_info["endCursor"]
        else:
            break

    return programs


def build_filename(program):
    """Build a stable, filesystem-safe filename that includes the program ID."""
    program_id = int(program["id"])
    slug = slugify(program["name"])
    return f"{program_id:03d}-{slug}.md"


def build_markdown(program):
    """Build the Markdown content for a program, covering one or more offers."""
    offers = [node["offer"] for node in program["programOffers"]["nodes"]]

    front_matter_lines = [
        "---",
        f"program_id: {program['id']}",
    ]
    if len(offers) == 1:
        offer = offers[0]
        front_matter_lines.append(f"offer_id: {offer['id']}")
        front_matter_lines.append(f"slug: \"{offer['slug']}\"")
    else:
        front_matter_lines.append("offers:")
        for offer in offers:
            front_matter_lines.append(f"  - offer_id: {offer['id']}")
            front_matter_lines.append(f"    slug: \"{offer['slug']}\"")
    front_matter_lines.append(f"color: \"{program['color']}\"")
    front_matter_lines.append("status: \"Ej påbörjad\"")
    front_matter_lines.append("ansvarig: \"\"")
    front_matter_lines.append("---")

    body_lines = ["", f"# {program['name']}", ""]

    if len(offers) == 1:
        description = offers[0]["description"] or ""
        body_lines.append("## Nuvarande text")
        body_lines.append("")
        body_lines.append(description)
        body_lines.append("")
        body_lines.append("## Ny text")
        body_lines.append("")
        body_lines.append("")
        body_lines.append("## Kommentarer")
        body_lines.append("")
    else:
        for offer in offers:
            description = offer["description"] or ""
            body_lines.append(f"## {offer['name']}")
            body_lines.append("")
            body_lines.append("### Nuvarande text")
            body_lines.append("")
            body_lines.append(description)
            body_lines.append("")
            body_lines.append("### Ny text")
            body_lines.append("")
            body_lines.append("")
            body_lines.append("### Kommentarer")
            body_lines.append("")

    return "\n".join(front_matter_lines + body_lines)


def main():
    print("Hämtar program från Expo...")
    programs = fetch_programs()
    print(f"✅ Hittade {len(programs)} program")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    created = 0
    skipped = 0

    for program in programs:
        filename = build_filename(program)
        filepath = os.path.join(OUTPUT_DIR, filename)

        if os.path.exists(filepath):
            print(f"⏭️ Hoppar över befintlig fil: {filename}")
            skipped += 1
            continue

        content = build_markdown(program)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(content)

        print(f"✅ Skapade: {filename}")
        created += 1

    print()
    print(f"Program från API: {len(programs)}")
    print(f"Skapade filer: {created}")
    print(f"Befintliga filer som hoppades över: {skipped}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"❌ Fel: {exc}", file=sys.stderr)
        sys.exit(1)
