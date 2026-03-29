"""
RIOS Main Orchestration
"""

from pathlib import Path
from datetime import datetime

from services.gmail_service import GmailService
from services.data_processor import DataProcessor
from config.settings import (
    WORK_EMAIL,
    OUTPUTS_DIR,
)
from utils.logger import logger


class RIOSAgent:
    """Main RIOS agent orchestrator."""

    def __init__(self):
        self.gmail = GmailService()
        self.processor = DataProcessor()

    def run(self):
        """Execute full RIOS workflow."""
        logger.info("=" * 60)
        logger.info("🧠 RIOS - Starting Weekly Report Processing")
        logger.info("=" * 60)

        try:
            logger.info("\n📧 Checking Gmail for new reports...")
            query = f'from:{WORK_EMAIL} subject:"Weekly Business Review" is:unread'
            messages = self.gmail.get_unread_messages(query=query)

            if not messages:
                logger.info("✅ No new reports found. RIOS is up to date.")
                return

            logger.info(f"✅ Found {len(messages)} new report(s)")

            for msg_idx, msg in enumerate(messages[:1], 1):
                logger.info(f"\n--- Processing Message {msg_idx} ---")

                message = self.gmail.get_message_details(msg["id"])
                sender = self.gmail.get_sender_email(message)
                subject = self.gmail.get_message_subject(message)

                logger.info(f"📧 From: {sender}")
                logger.info(f"📧 Subject: {subject}")

                logger.info("\n📥 Downloading attachments...")
                temp_dir = OUTPUTS_DIR / "temp_downloads"
                attachments = self.gmail.extract_attachments(message, temp_dir)

                if not attachments:
                    logger.warning("⚠️  No attachments found in message")
                    continue

                # Find ANY Excel file (.xlsx, .xls, .xlsm)
                excel_file = None
                for file_path in attachments:
                    suffix = str(file_path.suffix).lower()
                    if suffix in [".xlsx", ".xls", ".xlsm"]:
                        excel_file = file_path
                        logger.info(f"✅ Found Excel file: {excel_file.name}")
                        break

                if not excel_file:
                    logger.warning("⚠️  No Excel file found in attachments")
                    logger.info(f"Available files: {[f.name for f in attachments]}")
                    continue

                logger.info(f"✅ Using file: {excel_file.name}")

                logger.info("\n🔍 Analyzing retail data...")
                analysis = self.processor.analyze(str(excel_file))

                logger.info("\n💾 Saving reports...")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                one_pager_file = OUTPUTS_DIR / f"RIOS_OnePager_{timestamp}.txt"
                outreach_file = OUTPUTS_DIR / f"RIOS_VendorOutreach_{timestamp}.txt"

                with open(one_pager_file, "w", encoding="utf-8") as f:
                    f.write(analysis["one_pager"])

                with open(outreach_file, "w", encoding="utf-8") as f:
                    f.write(analysis["vendor_outreach"])

                logger.info(f"✅ Saved one-pager: {one_pager_file}")
                logger.info(f"✅ Saved vendor outreach: {outreach_file}")

                logger.info("\n📤 Sending results to work email...")
                email_body = f"""Hi Victor,

Your weekly RIOS analysis is ready.

Reports:
1. RIOS_OnePager.txt - Executive summary
2. RIOS_VendorOutreach.txt - Vendor talking points

Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

---
🧠 RIOS - Retail Intelligent Operating System
"""

                success = self.gmail.send_email(
                    to_email=WORK_EMAIL,
                    subject=f"✅ RIOS Weekly Analysis Ready",
                    body=email_body,
                    attachments=[one_pager_file, outreach_file],
                )

                if success:
                    self.gmail.mark_as_read(msg["id"])
                    logger.info("✅ Complete!")

        except Exception as e:
            logger.error(f"💥 Error: {e}", exc_info=True)

        logger.info("\n" + "=" * 60)
        logger.info("🧠 RIOS - Session Complete")
        logger.info("=" * 60)


if __name__ == "__main__":
    agent = RIOSAgent()
    agent.run()
