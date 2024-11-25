import discord
from discord.ext import commands
import pandas as pd
import difflib

# Load the data from an Excel file
file_path = r"C:\DiscordBot\values.xlsx.xlsx"  # Ensure the file is in the correct path
data = pd.read_excel(file_path)  # Load the data into a pandas DataFrame

# Debugging: Print the column names to verify the data structure
print("Columns in DataFrame:", data.columns)

# Bot setup: Enabling necessary intents and configuring the bot
intents = discord.Intents.default()
intents.message_content = True  # Allow the bot to read messages
bot = commands.Bot(command_prefix="!", intents=intents, case_insensitive=False)

# Clean up column names by stripping any extra spaces
data.columns = data.columns.str.strip()

# Utility function: Search for an item in the data (case-sensitive or case-insensitive)
def enhanced_find_item(data, item_name, case_sensitive=True):
    # Adjust for case sensitivity in item names
    item_name = item_name if case_sensitive else item_name.lower()
    data['Item name'] = data['Item name'].str.lower() if not case_sensitive else data['Item name']

    # First, try to find exact matches
    match = data[data['Item name'].str.contains(item_name, na=False)]
    if not match.empty:
        # If an exact match is found, return details
        row = match.iloc[0]
        return (
            f"**Item name**: {row['Item name']}\n"
            f"**Demand**: {row['Demand (out of 10)']}/10\n"
            f"**Value**: {row['Value']}\n"
            f"**Rate of change**: {row['rate of change']}"
        )
    else:
        # If no exact match is found, use fuzzy matching to suggest similar items
        item_names = data['Item name'].tolist()
        suggestions = difflib.get_close_matches(item_name, item_names, n=5, cutoff=0.5)
        if suggestions:
            # Provide suggestions if found
            suggestion_text = "\n".join([f"- {suggestion}" for suggestion in suggestions])
            return f"Item not found. Did you mean one of these?\n{suggestion_text}"
        else:
            return "Item not found. Please check the name and try again."

# Utility function: Find an exact or closest match using fuzzy matching
def find_exact_or_closest(item_name):
    # First, check for an exact match (case-insensitive)
    exact_match = data[data['Item name'].str.lower() == item_name.lower()]
    if not exact_match.empty:
        return exact_match.iloc[0]  # Return the first match

    # If no exact match, use fuzzy matching to find the closest match
    item_names = data['Item name'].str.lower().tolist()
    closest_match = difflib.get_close_matches(item_name.lower(), item_names, n=1, cutoff=0.6)
    if closest_match:
        return data[data['Item name'].str.lower() == closest_match[0]].iloc[0]
    return None  # Return None if no match is found

# Command: Show top X items based on a specific criterion (demand or value)
@bot.command(name="top")
async def top_items(ctx, number: int = 5, *, criterion: str = "demand"):
    criterion = criterion.lower()
    if criterion not in ["demand", "value"]:
        await ctx.send("Invalid criterion! Use `demand` or `value`.")
        return
    try:
        # Convert the relevant column to numeric and drop invalid entries
        column_name = "Demand (out of 10)" if criterion == "demand" else "Value"
        data[column_name] = pd.to_numeric(data[column_name], errors='coerce')
        sorted_data = data.dropna(subset=[column_name]).sort_values(by=column_name, ascending=False).head(number)

        # Create the response message for the top items
        response = "**Top Items:**\n" + "\n".join(
            [f"{row['Item name']} - {criterion.capitalize()}: {row[column_name]}" for _, row in sorted_data.iterrows()]
        )
        await ctx.send(response)
    except Exception as e:
        await ctx.send("An error occurred while fetching the top items.")
        print(e)

# Command: Filter items based on a given condition (supports Python-like syntax)
@bot.command(name="filter")
async def filter_items(ctx, *, condition: str):
    try:
        filtered_data = data.query(condition)  # Apply the filter to the data
        if filtered_data.empty:
            await ctx.send("No items match the given filter criteria.")
        else:
            # Create and send the filtered items list
            response = "**Filtered Items:**\n" + "\n".join(
                [f"{row['Item name']} - Demand: {row['Demand (out of 10)']} - Value: {row['Value']}" for _, row in filtered_data.iterrows()]
            )
            await ctx.send(response)
    except Exception as e:
        await ctx.send("Invalid filter criteria. Use Python-like syntax, e.g., `demand > 8`.")
        print(e)

# Command: Compare two items side by side (supports emojis as input)
@bot.command(name="compare")
async def compare(ctx, item1: str = None, item2: str = None):
    # Emoji to item name mapping for known emojis
    emoji_map = {
        "<:FestivalAura:1310704728765632644>": "Festival aura",  # Map the emoji to an item name
        "<:FrostSSJ4_Aura:1310698865510584411>": "FrostSSJ4 aura",  # Map the emoji to another item name
    }

    # Check if the input items are emojis and map them to item names
    if item1 in emoji_map:
        item1 = emoji_map[item1]
    if item2 in emoji_map:
        item2 = emoji_map[item2]
    
    print(f"Comparing: {item1} and {item2}")  # Debug log for comparison inputs
    
    # Ensure both items are provided
    if not item1 or not item2:
        await ctx.send("Please provide two items to compare. Example: `!compare item1 item2`")
        return

    # Get data for both items
    item1_data = find_exact_or_closest(item1)
    item2_data = find_exact_or_closest(item2)

    # Handle missing items
    if item1_data is None or item2_data is None:
        missing_items = []
        if item1_data is None:
            missing_items.append(item1)
        if item2_data is None:
            missing_items.append(item2)
        await ctx.send(f"One or both items not found: {', '.join(missing_items)}. Please check the names and try again.")
        return

    # Create and send the comparison response
    response = (
        f"**Comparison of {item1_data['Item name']} and {item2_data['Item name']}:**\n"
        f"**{item1_data['Item name']}:**\n"
        f"- Demand: {item1_data['Demand (out of 10)']}/10\n"
        f"- Value: {item1_data['Value']}\n"
        f"- Rate of Change: {item1_data['rate of change']}\n\n"
        f"**{item2_data['Item name']}:**\n"
        f"- Demand: {item2_data['Demand (out of 10)']}/10\n"
        f"- Value: {item2_data['Value']}\n"
        f"- Rate of Change: {item2_data['rate of change']}"
    )
    await ctx.send(response)

# Command: List recent updates based on the rate of change
@bot.command(name="recent")
async def recent_updates(ctx, number: int = 5):
    try:
        sorted_data = data.sort_values(by="rate of change", ascending=False).head(number)
        response = "**Recently Updated Items:**\n" + "\n".join(
            [f"{row['Item name']} - Rate of Change: {row['rate of change']}" for _, row in sorted_data.iterrows()]
        )
        await ctx.send(response)
    except Exception as e:
        await ctx.send("An error occurred while fetching recent updates.")
        print(e)

# Custom Help Command: Displays the list of available bot commands
bot.remove_command("help")  # Remove the default help command to replace with a custom one

@bot.command(name="help")
async def custom_help(ctx):
    help_text = """
**Available Commands:**
- `!value [item name]` or `!v [item name]`: Get the value, demand, and rate of change for an item.
- `!top [number] [criterion]`: List the top items based on `demand` or `value`. Example: `!top 5 demand`.
- `!filter [condition]`: Find items matching specific criteria. Example: `!filter demand > 8`.
- `!compare [item1] [item2]`: Compare two items side by side.
- `!recent [number]`: List items with the highest rate of change.
- `!help`: Show this help message.

**How to Use Commands:**
- Type the command with the required arguments.
- Example: `!value frost`
"""
    await ctx.send(help_text)

# Command: Get value, demand, and rate of change for a specific item
@bot.command(name="value", aliases=["v"])
async def value(ctx, *, item_name: str = None):
    if not item_name:
        await ctx.send("Please provide an item name. Example: `!value frost`")
        return

    # Emoji to item name mapping for known emojis
    emoji_map = {
        "<:FrostSSJ4_Aura:1310698865510584411>": "Frost aura",  # Emoji map example
        "<:FestivalAura:1310704728765632644>" : "Festival aura", 
    }

    # Check if the input is an emoji and convert it to the item name
    if item_name in emoji_map:
        item_name = emoji_map[item_name]

    # Find the item and send its details
    response = enhanced_find_item(data, item_name, case_sensitive=False)
    await ctx.send(response)

# Event: When the bot is ready, log the event
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user}")

# Event: Capture and process messages
@bot.event
async def on_message(message):
    if message.author.bot:
        return  # Ignore messages from other bots
    print(f"[{message.author}] {message.content}")  # Log messages for debugging
    await bot.process_commands(message)

# Run the bot with the provided token (replace 'token' with the actual bot token)
bot.run('token)  # Replace with your actual bot token
