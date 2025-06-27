from pymongo import MongoClient
from datetime import datetime

# MongoDB connection details
MONGO_URI = "mongodb://localhost:27017/"  # Replace with your MongoDB URI
DATABASE_NAME = "nova_db"  # Replace with your database name
COLLECTION_NAME = "conversations"  # Replace with your collection name

# Initialize MongoDB client and database
client = MongoClient(MONGO_URI)
db = client[DATABASE_NAME]
collection = db[COLLECTION_NAME]

def get_conversation_history(user_id):
    """
    Retrieve conversation history for a specific user from MongoDB.
    """
    try:
        history = list(collection.find({"user_id": user_id}, {"_id": 0}).sort("timestamp", 1))
        return history
    except Exception as e:
        print(f"Error fetching conversation history: {e}")
        return []

def update_conversation_history(user_id, user_input, nova_response):
    """
    Insert a new conversation entry into MongoDB.
    """
    try:
        conversation_entry = {
            "user_id": user_id,
            "user": user_input,
            "nova": nova_response,
            "timestamp": datetime.now()
        }
        collection.insert_one(conversation_entry)
    except Exception as e:
        print(f"Error updating conversation history: {e}")

def reset_conversation_history(user_id):
    """
    Delete all conversation history for a specific user from MongoDB.
    """
    try:
        collection.delete_many({"user_id": user_id})
    except Exception as e:
        print(f"Error resetting conversation history: {e}")