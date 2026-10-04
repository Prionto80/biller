from flask import Flask, request
from twilio.rest import Client
from twilio.twiml.messaging_response import MessagingResponse
import requests
import os
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Twilio credentials (from environment variables)
TWILIO_ACCOUNT_SID = os.getenv('TWILIO_ACCOUNT_SID')
TWILIO_AUTH_TOKEN = os.getenv('TWILIO_AUTH_TOKEN')
TWILIO_WHATSAPP_NUMBER = os.getenv('TWILIO_WHATSAPP_NUMBER')

# Local folder to save images (create if doesn't exist)
DOWNLOAD_FOLDER = os.getenv('DOWNLOAD_FOLDER', './downloads')
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)
    logger.info(f"Created download folder: {DOWNLOAD_FOLDER}")

client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

# Store last text message from each user
# Format: {phone_number: {'text': 'new orleans', 'timestamp': datetime}}
user_messages = {}


@app.route('/webhook', methods=['POST'])
def whatsapp_webhook():
    """
    Webhook endpoint for Twilio WhatsApp messages.
    - If text: Store it as the filename
    - If image: Download and save with stored filename
    """
    try:
        # Get incoming message data
        incoming_msg = request.values.get('Body', '').strip()
        sender_number = request.values.get('From')
        num_media = int(request.values.get('NumMedia', 0))
        
        logger.info(f"Received from {sender_number} | Text: '{incoming_msg}' | Media: {num_media}")
        
        # Initialize response
        resp = MessagingResponse()
        
        # Case 1: Text message received
        if incoming_msg and num_media == 0:
            # Store the text as filename for next image
            user_messages[sender_number] = {
                'text': incoming_msg,
                'timestamp': datetime.now()
            }
            logger.info(f"Stored filename for {sender_number}: {incoming_msg}")
            resp.message(f"✅ Got it! Next image will be saved as: `{incoming_msg}.jpg`")
            return str(resp)
        
        # Case 2: Image received
        if num_media > 0:
            # Check if we have a stored filename
            if sender_number not in user_messages:
                resp.message("❌ Please send me a name first! Example: 'new orleans'")
                logger.warning(f"No stored filename for {sender_number}")
                return str(resp)
            
            # Get the stored filename
            filename_without_ext = user_messages[sender_number]['text']
            filename = f"{filename_without_ext}.jpg"
            filepath = os.path.join(DOWNLOAD_FOLDER, filename)
            
            # Get image URL from Twilio
            media_url = request.values.get('MediaUrl0')
            media_type = request.values.get('MediaContentType0')
            
            logger.info(f"Downloading image from: {media_url}")
            
            try:
                # Download image from Twilio
                response = requests.get(media_url, timeout=30)
                response.raise_for_status()
                
                # Save image locally
                with open(filepath, 'wb') as f:
                    f.write(response.content)
                
                file_size = os.path.getsize(filepath) / 1024  # Size in KB
                logger.info(f"✅ Saved: {filepath} ({file_size:.1f} KB)")
                
                resp.message(f"✅ Downloaded! Saved as: `{filename}` ({file_size:.1f} KB)")
                
                # Clear the stored filename after use
                del user_messages[sender_number]
                
            except Exception as e:
                logger.error(f"Failed to download/save image: {str(e)}")
                resp.message(f"❌ Failed to save image: {str(e)}")
            
            return str(resp)
        
        # Case 3: Both text and image (unusual)
        if incoming_msg and num_media > 0:
            resp.message("Please send text and image separately!")
            return str(resp)
        
        # Case 4: Empty message
        resp.message("👋 Send me a name (text), then send an image!")
        return str(resp)
        
    except Exception as e:
        logger.error(f"Error in webhook: {str(e)}")
        resp = MessagingResponse()
        resp.message("⚠️ An error occurred. Please try again.")
        return str(resp)


@app.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint"""
    return {
        'status': 'ok',
        'download_folder': DOWNLOAD_FOLDER,
        'downloads': len(os.listdir(DOWNLOAD_FOLDER))
    }, 200


@app.route('/status', methods=['GET'])
def status():
    """Check bot status and saved files"""
    files = []
    if os.path.exists(DOWNLOAD_FOLDER):
        files = os.listdir(DOWNLOAD_FOLDER)
    
    return {
        'status': 'running',
        'download_folder': os.path.abspath(DOWNLOAD_FOLDER),
        'total_images': len(files),
        'files': sorted(files),
        'pending_filenames': list(user_messages.keys())
    }, 200


if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    logger.info(f"🤖 WhatsApp Bot Starting on port {port}")
    logger.info(f"📁 Saving images to: {os.path.abspath(DOWNLOAD_FOLDER)}")
    app.run(host='0.0.0.0', port=port, debug=False)
