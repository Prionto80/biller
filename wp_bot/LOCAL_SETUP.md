# WhatsApp Image Downloader - Local Setup (5 Minutes)

Download images from WhatsApp with custom names. Super simple!

## How It Works

```
You: "new orleans"
Bot: "✅ Got it! Next image will be saved as: new orleans.jpg"

You: [Send image]
Bot: "✅ Downloaded! Saved as: new orleans.jpg (45.2 KB)"
```

The image is automatically saved in your `./downloads` folder!

---

## Step 1: Install Python

Check if you have Python 3.8+:
```bash
python --version
```

If not, download from: https://www.python.org/downloads/

---

## Step 2: Create Twilio Account (2 minutes)

1. Go to https://www.twilio.com
2. Sign up for free
3. Go to **Console → WhatsApp**
4. Click **"Try WhatsApp"**
5. Note down these 3 things:
   - **Account SID** (starts with AC)
   - **Auth Token** (long string)
   - **Twilio WhatsApp Number** (like +1234567890)

6. Set up **WhatsApp Sandbox**:
   - Console → Messaging → WhatsApp → Sandbox Settings
   - You'll see a code like "join kitten-apple"
   - Send this code to Twilio's WhatsApp number
   - You'll get confirmation

---

## Step 3: Download & Prepare Bot

1. Download all files (especially `whatsapp_image_downloader.py` and `requirements.txt`)
2. Create a folder: `whatsapp-bot`
3. Put the files in that folder

---

## Step 4: Install Dependencies

Open terminal/command prompt in your `whatsapp-bot` folder:

```bash
pip install -r requirements.txt
```

---

## Step 5: Create .env File

In your `whatsapp-bot` folder, create a file named `.env` (no extension):

```
TWILIO_ACCOUNT_SID=ACxxxxxxxxxxxxx
TWILIO_AUTH_TOKEN=your_auth_token_here
TWILIO_WHATSAPP_NUMBER=+1234567890
DOWNLOAD_FOLDER=./downloads
PORT=5000
```

**Replace** the values with YOUR Twilio credentials!

---

## Step 6: Expose Your PC to Internet (ngrok)

Twilio needs to reach your computer. Install ngrok:

**Option A: Download**
- Go to https://ngrok.com/download
- Download for your OS
- Run it

**Option B: Terminal**
```bash
# macOS
brew install ngrok

# Windows (with Chocolatey)
choco install ngrok

# Linux
wget https://bin.equinox.io/c/...
```

---

## Step 7: Start ngrok

In terminal:
```bash
ngrok http 5000
```

You'll see output like:
```
Forwarding  https://abc123def456.ngrok.io -> http://localhost:5000
```

**Copy that URL!** (Keep ngrok running)

---

## Step 8: Start the Bot

In another terminal (in your `whatsapp-bot` folder):

```bash
python whatsapp_image_downloader.py
```

You should see:
```
🤖 WhatsApp Bot Starting on port 5000
📁 Saving images to: /Users/yourname/whatsapp-bot/downloads
```

**Keep this running!**

---

## Step 9: Tell Twilio Where Your Bot Is

1. Twilio Console → **Messaging → WhatsApp → Sandbox Settings**
2. Find "When a message comes in"
3. Change the URL to: `https://abc123def456.ngrok.io/webhook`
   (Use YOUR ngrok URL from Step 7)
4. Make sure it says **POST**
5. **SAVE**

---

## Step 10: Test It!

Send a WhatsApp message to Twilio's WhatsApp number:

**Message 1:** `new orleans`
- Bot replies: "✅ Got it! Next image will be saved as: new orleans.jpg"

**Message 2:** [Send an image]
- Bot replies: "✅ Downloaded! Saved as: new orleans.jpg (45.2 KB)"

Check your `./downloads` folder - you should see `new orleans.jpg`! 🎉

---

## More Examples

**Send multiple images:**

```
You: "sunset"
Bot: "✅ Got it! Next image will be saved as: sunset.jpg"
You: [Image 1]
Bot: "✅ Downloaded! Saved as: sunset.jpg"

You: "beach"
Bot: "✅ Got it! Next image will be saved as: beach.jpg"
You: [Image 2]
Bot: "✅ Downloaded! Saved as: beach.jpg"
```

Both images saved!

---

## Troubleshooting

### "ModuleNotFoundError: No module named 'flask'"
```bash
pip install -r requirements.txt
```

### Bot not responding
1. Check ngrok is still running
2. Check bot terminal shows no errors
3. Verify Twilio webhook URL is correct
4. Wait 30 seconds and try again

### ngrok connection keeps changing
When you restart ngrok, it gets a new URL. You need to:
1. Copy new URL from ngrok
2. Update Twilio webhook settings with new URL
3. Restart bot

### "Connection refused" error
Make sure bot is running! Start it with:
```bash
python whatsapp_image_downloader.py
```

### Images not saving
Check that `./downloads` folder exists:
```bash
mkdir downloads
```

---

## File Locations

- **Images saved to:** `./downloads/` (same folder as the bot)
- **Your PC's folder:** Right-click → Properties to see full path

---

## Check What's Saved

Open terminal in bot folder:

```bash
# macOS/Linux
ls downloads

# Windows
dir downloads
```

Or go to: File → Your whatsapp-bot folder → downloads folder

---

## Stop the Bot

Press `Ctrl + C` in the terminal

---

## Restart (Daily)

1. Start ngrok: `ngrok http 5000`
2. Start bot: `python whatsapp_image_downloader.py`
3. If ngrok URL changed, update Twilio webhook settings

---

## Advanced: Change Download Folder

Edit `.env`:
```
DOWNLOAD_FOLDER=/Users/yourname/Pictures/WhatsApp
```

Create that folder first!

---

## API Endpoints (For Testing)

**Check bot status:**
```
http://localhost:5000/status
```

Shows:
- Download folder location
- Total images saved
- List of all files

**Health check:**
```
http://localhost:5000/health
```

---

## Next Steps

✅ Create Twilio account  
✅ Set up WhatsApp sandbox  
✅ Download bot files  
✅ Install dependencies  
✅ Run ngrok + bot  
✅ Configure Twilio webhook  
✅ Send test message  
✅ Check downloads folder  
✅ Done! 🎉

Enjoy! Questions? Check the bot terminal for error messages.
