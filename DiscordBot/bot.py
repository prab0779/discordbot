import discord
from discord.ext import commands
import pandas as pd
import difflib

# Load the data from your file
file_path = r"C:\DiscordBot\values.xlsx.xlsx"  # Ensure the file is in the same directory
data = pd.read_excel(file_path)

# Inspect and print the column names to debug
print("Columns in DataFrame:", data.columns)

# Bot setup (case-sensitive commands with aliases)
intents = discord.Intents.default()
intents.message_content = True  # This allows the bot to read messages
bot = commands.Bot(command_prefix="!", intents=intents, case_insensitive=False)

# Ensure the columns are stripped of spaces
data.columns = data.columns.str.strip()

# Utility function to find items
def enhanced_find_item(data, item_name, case_sensitive=True):
    item_name = item_name if case_sensitive else item_name.lower()
    data['Item name'] = data['Item name'].str.lower() if not case_sensitive else data['Item name']

    # Check for exact matches
    match = data[data['Item name'].str.contains(item_name, na=False)]
    if not match.empty:
        row = match.iloc[0]
        return (
            f"**Item name**: {row['Item name']}\n"
            f"**Demand**: {row['Demand (out of 10)']}/10\n"
            f"**Value**: {row['Value']}\n"
            f"**Rate of change**: {row['rate of change']}"
        )
    else:
        # Use difflib to suggest close matches
        item_names = data['Item name'].tolist()
        suggestions = difflib.get_close_matches(item_name, item_names, n=5, cutoff=0.5)
        if suggestions:
            suggestion_text = "\n".join([f"- {suggestion}" for suggestion in suggestions])
            return f"Item not found. Did you mean one of these?\n{suggestion_text}"
        else:
            return "Item not found. Please check the name and try again."
        
def find_exact_or_closest(item_name):
    # Exact match (case-insensitive)
    exact_match = data[data['Item name'].str.lower() == item_name.lower()]
    if not exact_match.empty:
        return exact_match.iloc[0]  # Return the first match as a Series

    # Fuzzy match (suggest closest)
    item_names = data['Item name'].str.lower().tolist()
    closest_match = difflib.get_close_matches(item_name.lower(), item_names, n=1, cutoff=0.6)
    if closest_match:
        return data[data['Item name'].str.lower() == closest_match[0]].iloc[0]  # Return the closest match as a Series
    return None  # Return None if no match is found


# Command: Top X Items
@bot.command(name="top")
async def top_items(ctx, number: int = 5, *, criterion: str = "demand"):
    criterion = criterion.lower()
    if criterion not in ["demand", "value"]:
        await ctx.send("Invalid criterion! Use `demand` or `value`.")
        return
    try:
        sorted_data = data.sort_values(by=criterion.capitalize(), ascending=False).head(number)
        response = "**Top Items:**\n" + "\n".join(
            [f"{row['Item name']} - {criterion.capitalize()}: {row[criterion.capitalize()]}" for _, row in sorted_data.iterrows()]
        )
        await ctx.send(response)
    except Exception as e:
        await ctx.send("An error occurred while fetching the top items.")
        print(e)

# Command: Filter Items
@bot.command(name="filter")
async def filter_items(ctx, *, condition: str):
    try:
        filtered_data = data.query(condition)
        if filtered_data.empty:
            await ctx.send("No items match the given filter criteria.")
        else:
            response = "**Filtered Items:**\n" + "\n".join(
                [f"{row['Item name']} - Demand: {row['Demand (out of 10)']} - Value: {row['Value']}" for _, row in filtered_data.iterrows()]
            )
            await ctx.send(response)
    except Exception as e:
        await ctx.send("Invalid filter criteria. Use Python-like syntax, e.g., `demand > 8`.")
        print(e)

@bot.command(name="compare")
async def compare(ctx, item1: str, item2: str):
    # Find data for both items
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

    # Create and send comparison response
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



# Command: Recent Updates
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

# Command: Value (with autocomplete suggestions)
@bot.command(name="value", aliases=["v"])
async def value(ctx, *, item_name: str = None):
    if not item_name:
        await ctx.send("Please provide an item name. Example: `!value frost`")
        return
    response = enhanced_find_item(data, item_name, case_sensitive=False)  # Toggle case sensitivity here
    await ctx.send(response)

# Remove the default help command if you want a custom one
bot.remove_command("help")

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

@bot.event
async def on_ready():
    print(f"Logged in as {bot.user}")

@bot.event
async def on_message(message):
    if message.author.bot:
        return
    print(f"[{message.author}] {message.content}")
    await bot.process_commands(message)

# Run the bot
bot.run('token')  # Replace with your actual bot token
