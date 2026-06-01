# Bot Transport

Desktop automation app for exporting a JMS transportation report and sending a generated report image to a Feishu chat.

## Setup

1. Use Windows with Microsoft Excel installed. PNG generation uses Excel COM automation.
2. Install Python dependencies:

   ```bash
   pip install -r requirements.txt
   ```

3. Start the combined application:

   ```bash
   python Bot_Fei_Main.py
   ```

## Usage

- Configure scheduling, JMS export, and Feishu Chat ID credentials from the **Setting** tab.
- Select **Export JMS**, **Feishu Chat**, or both from the **Home** tab.
- When both workflows are selected, Export JMS runs first and Feishu Chat runs afterward.
