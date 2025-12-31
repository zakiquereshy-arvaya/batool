import asyncio
import os
import re
from dotenv import load_dotenv

from microsoft_teams.api import MessageActivity, TypingActivityInput
from microsoft_teams.apps import ActivityContext, App
from microsoft_teams.devtools import DevToolsPlugin
from microsoft.teams.ai import ChatPrompt
from microsoft.teams.mcpplugin import McpClientPlugin, McpClientPluginParams
from microsoft.teams.openai import OpenAICompletionsAIModel

# Load environment variables from .env file
load_dotenv()

# MCP Server Configuration - This is where all the tools live
mcp_server_api_key = os.getenv('FASTMCP_API_KEY')
mcp_server_url = "https://arvaya-availability.fastmcp.app/mcp"

# Azure OpenAI Configuration - Just for LLM inference
azure_model = os.getenv('AZURE_OPENAI_DEPLOYMENT')
azure_endpoint = os.getenv('AZURE_OPENAI_ENDPOINT', '').rstrip('/')

# Initialize Azure OpenAI model for ChatPrompt (LLM only - tools come from MCP)
# The MCP server handles all tool execution - this is just for LLM inference
azure_openai_model = OpenAICompletionsAIModel(
    model=azure_model,
    key=os.getenv('AZURE_OPENAI_API_KEY'),
    base_url=f"{azure_endpoint}/openai/deployments",
    api_version=os.getenv('AZURE_OPENAI_API_VERSION')
)

# The MCP server handles all tool execution and business logic
mcp_plugin = McpClientPlugin()
mcp_plugin.use_mcp_server(
    mcp_server_url,
    McpClientPluginParams(headers={
        "Authorization": f"Bearer {mcp_server_api_key}",
    })
)

# Create ChatPrompt - Thin wrapper that connects Azure OpenAI (LLM) + MCP Server (Tools)
# The MCP server does all the heavy lifting for tool execution
chat_prompt = ChatPrompt(
    azure_openai_model,
    plugins=[mcp_plugin]
)

app = App(plugins=[DevToolsPlugin()])


@app.on_message_pattern(re.compile(r"hello|hi|greetings"))
async def handle_greeting(ctx: ActivityContext[MessageActivity]) -> None:
    """Handle greeting messages."""
    await ctx.send("Hello! How can I assist you today?")


@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]):
    """
    Handle message activities - Thin client that delegates to MCP server.
    The MCP server handles all tool execution and business logic.
    """
    await ctx.reply(TypingActivityInput())

    try:
        user_message = ctx.activity.text
        
        # ChatPrompt automatically:
        # 1. Loads tools from MCP server
        # 2. Lets LLM decide when to use tools
        # 3. Executes tool calls via MCP server (MCP does the heavy lifting)
        # 4. Returns LLM response with tool results
        result = await chat_prompt.send(
            input=user_message,
            instructions="You are a helpful assistant in Microsoft Teams named Patti. Be concise and friendly. Use the available tools from the MCP server when needed to help answer questions."
        )
        
        if result.response and result.response.content:
            await ctx.reply(result.response.content)
        else:
            await ctx.reply("I'm sorry, I couldn't generate a response.")
    
    except Exception as e:
        error_msg = str(e)
        print(f"Error: {error_msg}")  # Debug logging
        await ctx.reply(f"Sorry, I encountered an error: {error_msg}")


def main():
    asyncio.run(app.start())


if __name__ == "__main__":
    main()