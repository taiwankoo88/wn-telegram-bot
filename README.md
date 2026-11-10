# WN Telegram bot

This Windows-hosted Telegram bot generates C3D1/C3D2 steel Written Notifications and Request Forms. The same bot handles fastener and gasket orders. This repository is a source handoff. Approved templates, production credentials, chat state, and submissions are supplied separately by the maintainer.

## Windows setup

1. Install Node.js and Python. Provide the existing authorized Codex runtime copy of the private Node dependency @oai/artifact-tool (source environment version 2.8.129), including its dependencies in node_modules. A standard npm install alone will not supply this private package.
2. Run python -m pip install -r wn/requirements.txt. For photo OCR, also install ocr/requirements.txt and copy table.example.json to table.json with absolute paths for the interpreter and script.
3. Copy steel.example.json to steel.json. Set root to the absolute path of the repository's wn folder and python to the interpreter installed above.
4. Place the approved templates in the paths below. Preserve filenames and formats. Keep business documents outside version control.

    wn/outputs/C3D1 Forms/SS_Steel Work Written Notification C3D1_Batch_2.docx
    wn/outputs/C3D1 Forms/SS_Request Form for Issuing Wirtten Notification_Batch_2 Amended.docx
    wn/outputs/C3D2 Forms/SS_Steel Work Written Notification C3D2.docx
    wn/outputs/C3D2 Forms/SS_Request Form.docx

5. For fastener orders and the full test suite, restore the approved templates/Gasket.xlsx, templates/HDG.xlsx and templates/SS316.xlsx. Desktop Word renders PDF previews and checks layout. Classic Outlook handles optional email when email.json is configured locally.
6. Run node test.mjs. With STEEL_SKIP_RENDER=1 it verifies the Word document contents without opening Word. With steel.json configured it generates both projects in a temporary folder and checks that a duplicate reference makes a revision.
7. Run Setup.cmd on the target Windows account to enter the authorized Telegram token and create the local DPAPI secret. Run Start.cmd after setup. Test with a separate Telegram bot token. When migrating production, stop the existing poller before starting the replacement and coordinate any transfer of authorized chat state.

See 使用说明.md for commands and the WN submission format. For Gemini Enterprise, grant its GitHub connection access to this private repository and ask it to read GEMINI.md and this README first. Deployment and live credentials are managed on the Windows host.
