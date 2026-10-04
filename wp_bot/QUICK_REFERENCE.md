# WhatsApp Image Downloader - Quick Reference

## What It Does

1. You send text: **"new orleans"**
2. You send image file
3. Bot saves it as: **new orleans.jpg** in your `downloads` folder

That's it! 🎯

---

## Quick Start (Checklist)

- [ ] Python 3.8+ installed
- [ ] Twilio account created
- [ ] Got 3 credentials (Account SID, Auth Token, WhatsApp Number)
- [ ] WhatsApp Sandbox set up
- [ ] Downloaded bot files
- [ ] Created `.env` file with credentials
- [ ] Ran `pip install -r requirements_local.txt`
- [ ] Started ngrok: `ngrok http 5000`
- [ ] Started bot: `python whatsapp_image_downloader.py`
- [ ] Updated Twilio webhook URL
- [ ] Tested by sending text + image

---

## Commands/Workflow

**CORRECT way to send images:**

```
Step 1: Send TEXT
You: "vacation photo"
Bot: "✅ Got it! Next image will be saved as: vacation photo.jpg"

Step 2: Send IMAGE
You: [Send your image file]
Bot: "✅ Downloaded! Saved as: vacation photo.jpg (120.5 KB)"

Result: File saved as "vacation photo.jpg"
```

**WRONG ways:**

❌ Send image without text first  
❌ Send text and image together  
❌ Send text only (without image after)

---

## File Structure

```
whatsapp-bot/
├── whatsapp_image_downloader.py  ← Main bot code
├── requirements_local.txt         ← Dependencies (pip install)
├── .env                          ← Your credentials (create this)
├── downloads/                    ← Images saved here
└── LOCAL_SETUP.md               ← Full setup instructions
```

---

## Three Terminals You'll Need

**Terminal 1: ngrok**
```bash
ngrok http 5000
# Keep running! Provides your public URL
```

**Terminal 2: Bot**
```bash
python whatsapp_image_downloader.py
# Keep running! Listens for messages
```

**Terminal 3: Checking stuff**
```bash
# List saved images
ls downloads

# Check bot status
curl http://localhost:5000/status
```

---

## How to Get Credentials

### Twilio Account SID & Auth Token
1. Go to https://console.twilio.com
2. You'll see it right on the dashboard
3. Copy and paste to `.env`

### Twilio WhatsApp Number
1. Console → Messaging → WhatsApp
2. Click "Sandbox Settings"
3. You'll see: "Your Sandbox Number"
4. Copy it (looks like +1234567890)

### ngrok URL
1. Run `ngrok http 5000`
2. Look for: `Forwarding  https://abc123def456.ngrok.io`
3. Copy the HTTPS URL
4. Put in Twilio webhook: `https://abc123def456.ngrok.io/webhook`

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "ModuleNotFoundError" | `pip install -r requirements_local.txt` |
| Bot not responding | Check ngrok running + Twilio webhook URL correct |
| Images not saving | Create `downloads` folder: `mkdir downloads` |
| "Connection refused" | Start bot: `python whatsapp_image_downloader.py` |
| Ngrok keeps disconnecting | Restart ngrok, update Twilio webhook with new URL |
| "TWILIO_ACCOUNT_SID not found" | Make sure `.env` file exists in same folder as bot |

---

## Bot Responses

| Scenario | Bot Says |
|----------|----------|
| You send text | "✅ Got it! Next image will be saved as: {text}.jpg" |
| You send image (with text) | "✅ Downloaded! Saved as: {text}.jpg (size KB)" |
| You send image (no text first) | "❌ Please send me a name first! Example: 'new orleans'" |
| You send empty message | "👋 Send me a name (text), then send an image!" |
| Error occurs | "⚠️ An error occurred. Please try again." |

---

## Bot Settings (In `.env`)

```
DOWNLOAD_FOLDER=./downloads
```
Change this to download to different location:
- `./downloads` (same folder as bot)
- `/Users/yourname/Pictures` (Mac)
- `C:\Users\yourname\Downloads` (Windows)
- `/home/username/Pictures` (Linux)

---

## Useful API Endpoints

**Check what's saved:**
```
http://localhost:5000/status
```

Response shows:
- Total images saved
- List of all files
- Download folder location

**Health check:**
```
http://localhost:5000/health
```

---

## Keep Bot Running 24/7

After you confirm it works:

### macOS/Linux:
Use `screen` or `tmux`
```bash
screen -S whatsapp
python whatsapp_image_downloader.py
# Press Ctrl+A then D to detach
```

### Windows:
Install as Windows Service or use Task Scheduler

### Better: Use cloud instead
If you want it running even when PC is off:
- Deploy to Render/Railway
- Modify bot to upload images to cloud storage

---

## Important Notes

⚠️ **Keep ngrok running** - If you close it, bot stops working  
⚠️ **Keep bot running** - If you close it, messages don't process  
⚠️ **ngrok URL changes** - When you restart ngrok, update Twilio webhook  
⚠️ **Free tier limited** - Twilio free tier has ~$15 credit/month

---

## Still Getting Errors?

1. Check bot terminal output (usually tells you what's wrong)
2. Check ngrok is running
3. Verify `.env` file has correct credentials
4. Try: `curl http://localhost:5000/health`
5. Check Twilio webhook URL in console

---

## Next Steps

1. ✅ Complete LOCAL_SETUP.md guide
2. ✅ Test with "new orleans" + image
3. ✅ Check downloads folder
4. ✅ Send more test images
5. ✅ Done! Start using! 🎉

---

**Happy downloading!** 📥
