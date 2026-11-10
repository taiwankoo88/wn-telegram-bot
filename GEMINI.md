# Maintainer handoff

Read README.md and the steel forms section of 使用说明.md first.

- bot.mjs routes Telegram messages and handles Telegram delivery. steel.mjs detects WN input, validates fields, invokes the updater, and saves a revision when a reference is reused.
- wn/outputs/update_steel_forms.py updates the two approved Word templates. steel_verify.py checks fields and signatures before delivery. RenderDocx.ps1 uses desktop Word for PDF and layout checks.
- Preserve both C3D1 and C3D2 behavior, Singapore dates, multiline descriptions, request form text boxes, signatures and job history. Test changes with node test.mjs before deployment. Use temporary output or a test bot for verification.
- Ask the maintainer for the approved templates and private runtime dependency before claiming an end-to-end test. Keep Telegram credentials, contact details, local configuration and production files out of commits.
