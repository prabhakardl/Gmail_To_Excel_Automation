import os
import json
import base64
import re
from email.utils import parsedate_to_datetime

import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


# ============================================================
# CONFIGURATION
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly"
]

USER_ID = "me"

MAX_MESSAGES = 100

# Gmail search query
#
# ""                       = all messages
# "in:inbox"               = inbox
# "is:unread"              = unread
# "has:attachment"         = attachments
# "from:example@gmail.com" = sender
# "subject:invoice"        = subject
#
GMAIL_QUERY = ""

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

OUTPUT_FOLDER = "output"
OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "gmail_messages.xlsx"
)


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate_gmail():

    print("[INFO] Checking Gmail authentication...")

    credentials = None

    # --------------------------------------------------------
    # Check credentials.json
    # --------------------------------------------------------

    if not os.path.exists(CREDENTIALS_FILE):

        print()
        print("[ERROR] credentials.json was not found.")
        print()
        print(
            "Place your Google OAuth Desktop credentials "
            "file beside main.py."
        )

        return None

    # --------------------------------------------------------
    # Validate credentials.json
    # --------------------------------------------------------

    try:

        with open(
            CREDENTIALS_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            json.load(file)

        print(
            "[INFO] credentials.json is valid JSON."
        )

    except Exception as error:

        print()
        print(
            "[ERROR] credentials.json is invalid."
        )
        print()
        print(error)

        return None

    # --------------------------------------------------------
    # Load token.json if available
    # --------------------------------------------------------

    if os.path.exists(TOKEN_FILE):

        print(
            "[INFO] token.json found."
        )

        try:

            # Check whether token.json is empty.
            if os.path.getsize(TOKEN_FILE) == 0:

                print(
                    "[WARNING] token.json is empty."
                )

                os.remove(TOKEN_FILE)

            else:

                # Validate JSON first.
                with open(
                    TOKEN_FILE,
                    "r",
                    encoding="utf-8"
                ) as file:

                    json.load(file)

                credentials = (
                    Credentials.from_authorized_user_file(
                        TOKEN_FILE,
                        SCOPES
                    )
                )

                print(
                    "[INFO] Existing Gmail token loaded."
                )

        except Exception as error:

            print(
                "[WARNING] token.json is invalid."
            )

            print(
                f"[WARNING] {error}"
            )

            print(
                "[INFO] Removing invalid token.json..."
            )

            try:

                os.remove(TOKEN_FILE)

            except Exception as remove_error:

                print(
                    "[ERROR] Could not remove token.json:"
                )

                print(remove_error)

                return None

            credentials = None

    # --------------------------------------------------------
    # Refresh existing credentials
    # --------------------------------------------------------

    if credentials:

        if credentials.valid:

            print(
                "[INFO] Existing Gmail authentication is valid."
            )

            return credentials

        if (
            credentials.expired
            and credentials.refresh_token
        ):

            print(
                "[INFO] Gmail token expired."
            )

            print(
                "[INFO] Refreshing token..."
            )

            try:

                credentials.refresh(
                    Request()
                )

                print(
                    "[INFO] Token refreshed successfully."
                )

                # Save refreshed token.
                with open(
                    TOKEN_FILE,
                    "w",
                    encoding="utf-8"
                ) as token_file:

                    token_file.write(
                        credentials.to_json()
                    )

                return credentials

            except Exception as error:

                print(
                    "[WARNING] Token refresh failed:"
                )

                print(error)

                credentials = None

    # --------------------------------------------------------
    # New OAuth authentication
    # --------------------------------------------------------

    print()
    print("=" * 70)
    print("STARTING GOOGLE OAUTH")
    print("=" * 70)
    print()
    print(
        "[INFO] Your browser will open for Google authentication."
    )
    print()

    try:

        flow = (
            InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE,
                SCOPES
            )
        )

        print(
            "[INFO] OAuth flow created successfully."
        )

        credentials = flow.run_local_server(
            port=0,
            open_browser=True
        )

        print()
        print(
            "[INFO] Google authentication completed."
        )

    except Exception as error:

        print()
        print("=" * 70)
        print("GOOGLE OAUTH ERROR")
        print("=" * 70)
        print()
        print(
            f"Error type: {type(error).__name__}"
        )
        print(
            f"Error message: {error}"
        )
        print()

        return None

    # --------------------------------------------------------
    # Save token
    # --------------------------------------------------------

    try:

        with open(
            TOKEN_FILE,
            "w",
            encoding="utf-8"
        ) as token_file:

            token_file.write(
                credentials.to_json()
            )

        print(
            "[INFO] New token saved to token.json."
        )

    except Exception as error:

        print(
            "[WARNING] Could not save token.json:"
        )

        print(error)

    return credentials


# ============================================================
# CREATE GMAIL SERVICE
# ============================================================

def create_gmail_service(credentials):

    print(
        "[INFO] Connecting to Gmail API..."
    )

    service = build(
        "gmail",
        "v1",
        credentials=credentials
    )

    print(
        "[INFO] Gmail API connection successful."
    )

    return service


# ============================================================
# GET HEADER
# ============================================================

def get_header(headers, header_name):

    header_name = header_name.lower()

    for header in headers:

        if (
            header.get("name", "").lower()
            == header_name
        ):

            return header.get(
                "value",
                ""
            )

    return ""


# ============================================================
# EMAIL ADDRESS
# ============================================================

def extract_email_address(value):

    if not value:
        return ""

    match = re.search(
        r"<([^<>@\s]+@[^<>@\s]+)>",
        value
    )

    if match:

        return match.group(1)

    match = re.search(
        r"[A-Za-z0-9._%+-]+@"
        r"[A-Za-z0-9.-]+\.[A-Za-z]{2,}",
        value
    )

    if match:

        return match.group(0)

    return ""


# ============================================================
# BASE64 DECODE
# ============================================================

def decode_base64(data):

    if not data:
        return ""

    try:

        decoded = base64.urlsafe_b64decode(
            data + "=" * (-len(data) % 4)
        )

        return decoded.decode(
            "utf-8",
            errors="replace"
        )

    except Exception:

        return ""


# ============================================================
# HTML TO TEXT
# ============================================================

def html_to_text(content):

    if not content:
        return ""

    content = re.sub(
        r"<script.*?</script>",
        "",
        content,
        flags=re.DOTALL | re.IGNORECASE
    )

    content = re.sub(
        r"<style.*?</style>",
        "",
        content,
        flags=re.DOTALL | re.IGNORECASE
    )

    content = re.sub(
        r"<br\s*/?>",
        "\n",
        content,
        flags=re.IGNORECASE
    )

    content = re.sub(
        r"</p\s*>",
        "\n",
        content,
        flags=re.IGNORECASE
    )

    content = re.sub(
        r"<[^>]+>",
        "",
        content
    )

    return content.strip()


# ============================================================
# EXTRACT EMAIL BODY
# ============================================================

def extract_body(payload):

    if not payload:
        return ""

    body = payload.get(
        "body",
        {}
    )

    data = body.get(
        "data"
    )

    mime_type = payload.get(
        "mimeType",
        ""
    )

    if data:

        content = decode_base64(
            data
        )

        if mime_type == "text/html":

            return html_to_text(
                content
            )

        return content.strip()

    plain_text = ""
    html_text = ""

    for part in payload.get(
        "parts",
        []
    ):

        part_mime = part.get(
            "mimeType",
            ""
        )

        part_body = part.get(
            "body",
            {}
        )

        part_data = part_body.get(
            "data"
        )

        if part_data:

            content = decode_base64(
                part_data
            )

            if part_mime == "text/plain":

                if not plain_text:

                    plain_text = content

            elif part_mime == "text/html":

                if not html_text:

                    html_text = content

        nested = part.get(
            "parts",
            []
        )

        if nested:

            nested_text = extract_body(
                {
                    "parts": nested
                }
            )

            if nested_text and not plain_text:

                plain_text = nested_text

    if plain_text:

        return plain_text.strip()

    if html_text:

        return html_to_text(
            html_text
        )

    return ""


# ============================================================
# ATTACHMENTS
# ============================================================

def extract_attachments(payload):

    names = []

    if not payload:
        return names

    for part in payload.get(
        "parts",
        []
    ):

        filename = part.get(
            "filename",
            ""
        )

        attachment_id = (
            part.get(
                "body",
                {}
            ).get(
                "attachmentId"
            )
        )

        if filename and attachment_id:

            names.append(
                filename
            )

        nested = part.get(
            "parts",
            []
        )

        if nested:

            names.extend(
                extract_attachments(
                    {
                        "parts": nested
                    }
                )
            )

    return names


# ============================================================
# FORMAT DATE
# ============================================================

def format_date(value):

    if not value:
        return ""

    try:

        date_value = parsedate_to_datetime(
            value
        )

        return date_value.strftime(
            "%Y-%m-%d %H:%M:%S %z"
        )

    except Exception:

        return value


# ============================================================
# GET MESSAGE IDS
# ============================================================

def get_message_ids(
    service,
    query,
    maximum
):

    print()
    print(
        "[INFO] Searching Gmail..."
    )

    print(
        f"[INFO] Query: {query or 'ALL MESSAGES'}"
    )

    print(
        f"[INFO] Maximum messages: {maximum}"
    )

    message_ids = []

    page_token = None

    while len(message_ids) < maximum:

        remaining = (
            maximum
            - len(message_ids)
        )

        response = (
            service.users()
            .messages()
            .list(
                userId=USER_ID,
                q=query,
                maxResults=min(
                    100,
                    remaining
                ),
                pageToken=page_token
            )
            .execute()
        )

        messages = response.get(
            "messages",
            []
        )

        for message in messages:

            message_ids.append(
                message["id"]
            )

            if len(message_ids) >= maximum:

                break

        page_token = response.get(
            "nextPageToken"
        )

        if not page_token:

            break

    print(
        f"[INFO] Messages found: "
        f"{len(message_ids)}"
    )

    return message_ids


# ============================================================
# GET FULL MESSAGE
# ============================================================

def get_full_message(
    service,
    message_id
):

    return (
        service.users()
        .messages()
        .get(
            userId=USER_ID,
            id=message_id,
            format="full"
        )
        .execute()
    )


# ============================================================
# CONVERT MESSAGE
# ============================================================

def convert_message(message):

    payload = message.get(
        "payload",
        {}
    )

    headers = payload.get(
        "headers",
        []
    )

    sender = get_header(
        headers,
        "From"
    )

    attachments = extract_attachments(
        payload
    )

    return {

        "Message_ID":
            message.get(
                "id",
                ""
            ),

        "Thread_ID":
            message.get(
                "threadId",
                ""
            ),

        "Date":
            format_date(
                get_header(
                    headers,
                    "Date"
                )
            ),

        "From":
            sender,

        "From_Email":
            extract_email_address(
                sender
            ),

        "To":
            get_header(
                headers,
                "To"
            ),

        "Cc":
            get_header(
                headers,
                "Cc"
            ),

        "Bcc":
            get_header(
                headers,
                "Bcc"
            ),

        "Subject":
            get_header(
                headers,
                "Subject"
            ),

        "Labels":
            ", ".join(
                message.get(
                    "labelIds",
                    []
                )
            ),

        "Has_Attachment":
            "Yes"
            if attachments
            else "No",

        "Attachment_Names":
            ", ".join(
                attachments
            ),

        "Body":
            extract_body(
                payload
            )
    }


# ============================================================
# EXPORT EXCEL
# ============================================================

def export_excel(records):

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    dataframe = pd.DataFrame(
        records
    )

    summary = pd.DataFrame({

        "Metric": [

            "Total Messages",

            "Messages With Attachments",

            "Unique Senders",

            "Unique Subjects"
        ],

        "Value": [

            len(records),

            sum(
                1
                for row in records
                if row["Has_Attachment"]
                == "Yes"
            ),

            dataframe[
                "From_Email"
            ].nunique(),

            dataframe[
                "Subject"
            ].nunique()
        ]
    })

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        dataframe.to_excel(
            writer,
            sheet_name="Gmail_Messages",
            index=False
        )

        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

    print()
    print(
        "[INFO] Excel file created:"
    )

    print(
        os.path.abspath(
            OUTPUT_FILE
        )
    )

    print(
        f"[INFO] Records exported: "
        f"{len(records)}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 70)
    print(
        "        GMAIL → EXCEL AUTOMATION"
    )
    print("=" * 70)
    print()

    credentials = authenticate_gmail()

    if not credentials:

        print()
        print(
            "[ERROR] Gmail authentication failed."
        )

        return

    try:

        service = create_gmail_service(
            credentials
        )

    except Exception as error:

        print()
        print(
            "[ERROR] Could not connect to Gmail API."
        )

        print(
            f"{type(error).__name__}: {error}"
        )

        return

    try:

        message_ids = get_message_ids(
            service,
            GMAIL_QUERY,
            MAX_MESSAGES
        )

    except Exception as error:

        print()
        print(
            "[ERROR] Could not read Gmail messages."
        )

        print(
            f"{type(error).__name__}: {error}"
        )

        return

    if not message_ids:

        print(
            "[INFO] No messages found."
        )

        return

    records = []

    print()
    print(
        "[INFO] Downloading messages..."
    )
    print()

    for number, message_id in enumerate(
        message_ids,
        start=1
    ):

        try:

            message = get_full_message(
                service,
                message_id
            )

            record = convert_message(
                message
            )

            records.append(
                record
            )

            subject = (
                record["Subject"]
                or "(No Subject)"
            )

            print(
                f"[{number}/{len(message_ids)}] "
                f"{subject[:70]}"
            )

        except Exception as error:

            print(
                f"[ERROR] Message "
                f"{message_id}: "
                f"{error}"
            )

    if records:

        export_excel(
            records
        )

        print()
        print("=" * 70)
        print(
            "             COMPLETED"
        )
        print("=" * 70)

    else:

        print(
            "[ERROR] No messages were exported."
        )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    main()
