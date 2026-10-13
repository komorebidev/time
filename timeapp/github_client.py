import os

import requests
from dotenv import load_dotenv

load_dotenv()

OWNER = "komorebidev"
REPOSITORY = "time"


def get_monthly_log(year: int, month: int) -> str:
    token = os.getenv("GITHUB_TOKEN")

    if not token:
        raise RuntimeError("GITHUB_TOKEN is not configured.")

    path = f"{year}/{year}-{month:02d}.md"
    url = (
        f"https://api.github.com/repos/"
        f"{OWNER}/{REPOSITORY}/contents/{path}"
    )

    response = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github.raw+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        timeout=20,
    )

    response.raise_for_status()
    return response.text
