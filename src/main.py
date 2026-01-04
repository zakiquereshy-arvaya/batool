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
calendar_mcp_server_url = "https://arvaya-availability.fastmcp.app/mcp"
time_entry_mcp_server_url= "https://billi-tool.fastmcp.app/mcp"

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
mcp_plugin_booking = McpClientPlugin()
mcp_plugin_booking.use_mcp_server(
    calendar_mcp_server_url,
    McpClientPluginParams(headers={
        "Authorization": f"Bearer {mcp_server_api_key}",
    })
)

mcp_plugin_time_entry = McpClientPlugin()
mcp_plugin_time_entry.use_mcp_server(
    time_entry_mcp_server_url,
    McpClientPluginParams(headers={
        "Authorization": f"Bearer {mcp_server_api_key}",
    })
)

# Create ChatPrompt - Thin wrapper that connects Azure OpenAI (LLM) + MCP Server (Tools)
# The MCP server does all the heavy lifting for tool execution
chat_prompt = ChatPrompt(
    azure_openai_model,
    plugins=[mcp_plugin_booking, mcp_plugin_time_entry]
)

# Store conversation history per conversation thread
# This preserves context that the MCP server needs
conversation_history: dict[str, list] = {}

app = App(plugins=[DevToolsPlugin()])

"""
@app.on_message_pattern(re.compile(r"hello|hi|greetings"))
async def handle_greeting(ctx: ActivityContext[MessageActivity]) -> None:
    await ctx.send("Hello! How can I assist you today?")

"""

@app.on_message
async def handle_message(ctx: ActivityContext[MessageActivity]):
    """
    Handle message activities - Thin client that delegates to MCP server.
    The MCP server handles all tool execution and business logic.
    Maintains conversation history to preserve context.
    """
    await ctx.reply(TypingActivityInput())

    try:
        user_message = ctx.activity.text
        conversation_id = ctx.activity.conversation.id
        
        # Extract user metadata from the message activity
        # This provides context about who is making the request
        user_info = {"name": "", "id": "", "aad_object_id": ""}
        
        # Try to access user information from the activity
        # Microsoft Teams activities have 'from_property' or 'from' attribute with user info
        # Use getattr to avoid conflicts with Python's 'from' keyword
        user_from = None
        if hasattr(ctx.activity, 'from_property'):
            user_from = getattr(ctx.activity, 'from_property', None)
        if not user_from and hasattr(ctx.activity, 'from'):
            user_from = getattr(ctx.activity, 'from', None)
        
        if user_from:
            # Extract user properties safely
            if hasattr(user_from, 'name'):
                user_info["name"] = user_from.name or ""
            if hasattr(user_from, 'id'):
                user_info["id"] = user_from.id or ""
            if hasattr(user_from, 'aad_object_id'):
                user_info["aad_object_id"] = user_from.aad_object_id or ""
            elif hasattr(user_from, 'aadObjectId'):
                user_info["aad_object_id"] = user_from.aadObjectId or ""
        
        # Debug: Print user info to verify extraction (remove in production if needed)
        if user_info.get("name"):
            print(f"User metadata extracted: {user_info}")
        
        # Initialize conversation history if this is a new conversation
        if conversation_id not in conversation_history:
            conversation_history[conversation_id] = []
        
        # Build conversation context: include previous messages so MCP server has full context
        # This preserves context that the MCP server needs (like dates, names, etc.)
        # The MCP server needs to see the full conversation to maintain context
        # Include user metadata naturally in the context for tools that need userName/sender
        user_name = user_info.get('name', '')
        user_context_prefix = f"User ({user_name})" if user_name else "User"
        
        # Build contextual input with user metadata
        if len(conversation_history[conversation_id]) > 0:
            # Include recent conversation history for context (last 5 exchanges)
            recent_history = conversation_history[conversation_id][-5:]  # Last 5 messages for context
            context_summary = "\n".join([
                f"User: {msg['content']}" if msg['role'] == 'user' else f"Assistant: {msg['content']}"
                for msg in recent_history
            ])
            contextual_input = f"Conversation history:\n{context_summary}\n\n{user_context_prefix} request: {user_message}"
        else:
            contextual_input = f"{user_context_prefix} request: {user_message}"
        
        # Add user metadata to instructions dynamically if available
        user_metadata_note = ""
        if user_name:
            user_metadata_note = f"\n\n## CURRENT USER CONTEXT\nThe person making this request is: {user_name}. Use this name for tools that require userName or sender parameters."
        
        # Build instructions with user metadata if available
        base_instructions = """CRITICAL SYSTEM REQUIREMENT
You MUST NOT use emojis, special characters, or any Unicode characters above ASCII 255.
Use only plain text: letters (A-Z, a-z), numbers (0-9), and basic punctuation (. , ! ? - ' ").
Violations will cause system errors. This is a hard technical constraint.

You are Billi, a friendly and efficient assistant for ARVAYA Consulting.

## YOUR ROLE
You are a THIN CLIENT that routes user requests to MCP server tools. The MCP servers handle ALL intelligence, date parsing, business logic, and data processing. You do NOT interpret, convert, or process dates, times, or any data - you ONLY call tools with the exact information from the user.

## YOUR AVAILABLE TOOLS

### Time Entry Tool (from time entry MCP server):
- process_time_entry - Submit time entries to QuickBooks and Monday.com
  When to use: User mentions logging time, time entry, hours worked, submitting time
  Required: messageText (use the EXACT user message), userName

### Calendar Tools (from calendar MCP server):
- get_users_with_name_and_email - Find user email addresses by name
  When to use: User mentions a person's name and you need their email address
  
- check_availability - Check calendar availability for a user
  When to use: User asks about availability, free time, when someone is available, schedule
  Required: user_email (get from get_users_with_name_and_email first if only name provided)
  IMPORTANT: Pass dates/times EXACTLY as the user said them - "tomorrow", "next Monday", "1/3/2026", etc. The MCP server will parse them.
  
- book_meeting - Book a meeting in a user's calendar
  When to use: User wants to schedule, book, or set up a meeting, appointment
  Required: user_email, subject, start_datetime, end_datetime, sender
  IMPORTANT: Pass dates/times EXACTLY as the user said them. DO NOT convert "tomorrow" to a date - pass "tomorrow" to the MCP server.

## USER CONTEXT
The user making this request is identified in the message. Use their name when tools require a sender or userName parameter.

## CRITICAL INSTRUCTIONS - READ CAREFULLY

1. DO NOT interpret or convert dates/times. If user says "tomorrow", pass "tomorrow" to MCP tools. If user says "1/3/2026", pass "1/3/2026". The MCP server handles ALL date parsing.

2. DO NOT add your own date interpretations to responses. Only use dates that come back from MCP tool results.

3. When a user asks about availability or booking:
   - DO NOT just greet them - IMMEDIATELY use the appropriate tool
   - If they mention a person's name, call get_users_with_name_and_email FIRST to get their email
   - Then call check_availability or book_meeting with the email address
   - Pass dates/times EXACTLY as the user provided them - the MCP server will parse them correctly

4. For book_meeting, the sender parameter should be the name of the person making the request (from the user context).
5. For process_time_entry, the userName parameter should be the name of the person making the request (from the user context).

6. Tool Usage Examples:
   - User: "check availability of David Hogg" -> call get_users_with_name_and_email("David Hogg"), then check_availability with returned email
   - User: "book a meeting with Sarah tomorrow at 2pm" -> call get_users_with_name_and_email("Sarah"), then book_meeting with start_datetime="tomorrow 2pm" (NOT a converted date)
   - User: "book Ryan for 1/3/2026 at 9am" -> call get_users_with_name_and_email("Ryan"), then book_meeting with start_datetime="1/3/2026 9am"
   - User: "who is John Smith" -> call get_users_with_name_and_email("John Smith")
   - User: "log 4 hours for project X" -> call process_time_entry with messageText="log 4 hours for project X"

## DATE/TIME HANDLING - CRITICAL
- The MCP servers handle ALL date parsing, timezone conversion, and business logic
- You MUST pass dates/times EXACTLY as the user provides them: "tomorrow", "today", "next Monday", "1/3/2026", "2pm", etc.
- DO NOT convert "tomorrow" to "October 27th" or any specific date - pass "tomorrow" to the MCP server
- DO NOT interpret relative dates - let the MCP server do it
- Only use specific dates in your responses if they come from MCP tool results

## COMMUNICATION STYLE
- First interaction: Brief greeting only if no action is requested
- When user requests an action: Execute the tool IMMEDIATELY with exact user input, then confirm the result
- Be concise and action-oriented
- After tool execution, provide a clear summary using information from the tool results"""
        
        # Append user metadata note if available
        full_instructions = base_instructions + user_metadata_note
        
        result = await chat_prompt.send(
            input=contextual_input,  # Include conversation history for MCP server context
            instructions=full_instructions
        )
        
        if result.response and result.response.content:
            ai_response = result.response.content
            # Add both user message and AI response to conversation history for next iteration
            conversation_history[conversation_id].append({
                "role": "user",
                "content": user_message
            })
            conversation_history[conversation_id].append({
                "role": "assistant",
                "content": ai_response
            })
            await ctx.reply(ai_response)
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