from flask import Flask, jsonify, request
from openai import OpenAI
import os
import time

app = Flask(__name__)
client = OpenAI(
    api_key=os.environ['OPENAI_API_KEY'],
)
# Define assistant IDs
assistant_ids = {
    "extract_keywords": "asst_Tz2MjPx96CAYUadc3242uc90",
    "mass_morph": "asst_Tz2MjPx96CAYUadc3242uc90",
    "strategic_quant_points": "asst_r5jX4HHgPDdnTBUjgZU08xpp",
    "clearly_evident_quant_points": "asst_eaJzE6iYlOMvICYQQFqfgbQQ"
}

def retrieve_assistant(assistant_id):
    try:
        assistant = client.beta.assistants.retrieve(assistant_id)
        return assistant
    except Exception as e:
        return {"error": str(e)}

@app.route('/assistant1', methods=['GET'])
def get_assistant1():
    assistant = retrieve_assistant(assistant_ids["assistant1"])
    return jsonify(assistant)

@app.route('/assistant2', methods=['GET'])
def get_assistant2():
    assistant = retrieve_assistant(assistant_ids["assistant2"])
    return jsonify(assistant)

@app.route('/assistant3', methods=['GET'])
def get_assistant3():
    assistant = retrieve_assistant(assistant_ids["assistant3"])
    return jsonify(assistant)

@app.route('/assistant4', methods=['GET'])
def get_assistant4():
    assistant = retrieve_assistant(assistant_ids["assistant4"])
    return jsonify(assistant)

@app.route('/assistant5', methods=['GET'])
def get_assistant5():
    assistant = retrieve_assistant(assistant_ids["assistant5"])
    return jsonify(assistant)

@app.route('/create_thread_and_run', methods=['POST'])
def create_thread_and_run():
    user_message = request.json.get('message', 'Create thread and run')
    try:
        run = client.beta.threads.create_and_run(
            assistant_id=assistant_ids["extract_keywords"],
            thread={
                "messages": [
                    {"role": "user", "content": user_message}
                ]
            },
            temperature=0  # Set temperature to 0
        )
        
        # Retrieve the thread ID from the run object
        thread_id = run.thread_id
        
        # Wait until there are more than one message in the thread and all messages have non-empty content
        messages_list = []
        while len(messages_list) < 2 or any(not msg.content for msg in messages_list):
            messages = client.beta.threads.messages.list(thread_id=thread_id, order='desc')
            messages_list = list(messages)
            if len(messages_list) < 2 or any(not msg.content for msg in messages_list):
                time.sleep(1)  # Wait for 1 second before checking again
        
        # Get the most recent message ID
        message_id = messages_list[0].id if messages_list else None
        
        if message_id:
            message = client.beta.threads.messages.retrieve(
                message_id=message_id,
                thread_id=thread_id,
            )
            return jsonify({
                "run": run.to_dict(),  # Convert the run object to a dictionary
                "message": message.to_dict()  # Convert the message object to a dictionary
            })
        else:
            return jsonify({"error": "No messages found in the thread"}), 404
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/retrieve_message', methods=['GET'])
def retrieve_message():
    thread_id = request.args.get('thread_id')
    message_id = request.args.get('message_id')
    if not thread_id or not message_id:
        return jsonify({"error": "thread_id and message_id are required"}), 400

    try:
        message = client.beta.threads.messages.retrieve(
            message_id=message_id,
            thread_id=thread_id,
        )
        return jsonify(message.to_dict())  # Convert the message object to a dictionary
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)
    app.run(debug=True)