import asyncio
import os
import re
from dotenv import load_dotenv
from openai import AzureOpenAI


from microsoft_teams.api import MessageActivity, TypingActivityInput
from microsoft_teams.apps import ActivityContext, App
from microsoft_teams.devtools import DevToolsPlugin

# Load environment variables from .env file
load_dotenv()

azure_model = os.getenv('AZURE_OPENAI_DEPLOYMENT')

azure_openai_client = AzureOpenAI(
    api_version = os.getenv('AZURE_OPENAI_API_VERSION'),
    azure_endpoint = os.getenv('AZURE_OPENAI_ENDPOINT'),
    api_key = os.getenv('AZURE_OPENAI_API_KEY')
)

app = App(plugins=[DevToolsPlugin()])



@app.on_message_pattern(re.compile(r"hello|hi|greetings"))
async def handle_greeting(ctx: ActivityContext[MessageActivity]) -> None:
    """Handle greeting messages."""
    await ctx.send("Hello! How can I assist you today?")


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]):
    """Handle message activities using the new generated handler system."""
    await ctx.reply(TypingActivityInput())

    try:
        user_message = ctx.activity.text
        response = azure_openai_client.chat.completions.create(
           
            messages=[
                {
                    "role": "system",
                    "content": "You are a helpful assistant in microsoft teams named Patti. Be concise and friendly"
                },
                {
                    "role": "user",
                    "content": user_message
                }
            ],
            max_tokens=16384,
            temperature=0.8,
            model=azure_model 
        )
        ai_response = response.choices[0].message.content
        
        await ctx.reply(ai_response)
    
    except Exception as e:
        await ctx.reply(f"Sorry, I encountered an error: {str(e)}")

def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()
