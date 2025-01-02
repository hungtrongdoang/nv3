import requests

# Token for the bot
BOT_TOKEN = "7006593154:AAEejM7r2sOeFOV-FzIIb-7ek5gZl1A2D3M"
BASE_URL = f"https://api.telegram.org/bot{BOT_TOKEN}"

# Function to send a test message
def send_test_message():
    try:
        # Retrieve updates to find the latest chat ID
        updates_url = f"{BASE_URL}/getUpdates"
        updates_response = requests.get(updates_url)
        updates_data = updates_response.json()

        # Ensure there are updates and extract the latest chat ID
        if updates_data.get("result"):
            chat_id = updates_data["result"][-1]["message"]["chat"]["id"]

            # Send a test message
            message_url = f"{BASE_URL}/sendMessage"
            message_payload = {"chat_id": chat_id, "text": "This is a test message from your bot!"}
            message_response = requests.post(message_url, json=message_payload)

            if message_response.status_code == 200:
                print("Test message sent successfully!")
                print("Response:", message_response.json())
            else:
                print(f"Failed to send message. Status code: {message_response.status_code}")
                print("Response:", message_response.json())
        else:
            print("No updates found. Ensure you have sent a message to your bot first.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Test the bot
send_test_message()
