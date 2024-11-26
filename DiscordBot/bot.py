import discord
from discord.ext import commands
import pandas as pd
import difflib
import os
import openpyxl
from discord import Embed
import re

# Load the data from an Excel file
file_path = r"values.xlsx.xlsx"  # Ensure the file is in the correct path
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
    data['Item name'] = data['Item name'].str.lower(
    ) if not case_sensitive else data['Item name']

    # First, try to find exact matches
    match = data[data['Item name'].str.contains(item_name, na=False)]
    if not match.empty:
        # If an exact match is found, return details
        row = match.iloc[0]
        return (f"**Item name**: {row['Item name']}\n"
                f"**Demand**: {row['Demand (out of 10)']}/10\n"
                f"**Value**: {row['Value']}\n"
                f"**Rate of change**: {row['rate of change']}")
    else:
        # If no exact match is found, use fuzzy matching to suggest similar items
        item_names = data['Item name'].tolist()
        suggestions = difflib.get_close_matches(item_name,
                                                item_names,
                                                n=5,
                                                cutoff=0.5)
        if suggestions:
            # Provide suggestions if found
            suggestion_text = "\n".join(
                [f"- {suggestion}" for suggestion in suggestions])
            return f"Item not found. Did you mean one of these?\n{suggestion_text}"
        else:
            return "Item not found. Please check the name and try again."


import re


# Utility function: Find an exact or closest match using fuzzy matching
def find_exact_or_closest(item_name):
    """
    Search for the item in the Excel sheet by exact match or closest match.
    Assumes 'data' is a DataFrame containing item data.
    """
    # Remove emojis and extra spaces from item names
    sanitized_name = re.sub(r"<:[^:]+:[0-9]+>", "", item_name).strip()

    # Ensure 'data' exists and contains the necessary columns
    if 'Item name' not in data.columns or 'Value' not in data.columns:
        return None

    # Try to find an exact match
    match = data[data['Item name'].str.casefold() == sanitized_name.casefold()]
    if not match.empty:
        return match.iloc[0]

    # Fallback: Find closest match (case-insensitive substring match)
    closest_match = data[data['Item name'].str.contains(sanitized_name,
                                                        case=False)]
    if not closest_match.empty:
        return closest_match.iloc[0]

    # If no match is found, return None
    return None


# Command: Show top X items based on a specific criterion (demand or value)
@bot.command(name="top")
async def top_items(ctx, number: int = 5, *, criterion: str = "demand"):
    criterion = criterion.lower()
    if criterion not in ["demand", "value"]:
        await ctx.reply("Invalid criterion! Use `demand` or `value`.",
                        mention_author=False)
        return
    try:
        # Convert the relevant column to numeric and drop invalid entries
        column_name = "Demand (out of 10)" if criterion == "demand" else "Value"
        data[column_name] = pd.to_numeric(data[column_name], errors='coerce')
        sorted_data = data.dropna(subset=[column_name]).sort_values(
            by=column_name, ascending=False).head(number)

        if sorted_data.empty:
            await ctx.reply(
                f"No items found for the specified criterion: {criterion}.",
                mention_author=False)
            return

        # Create an embed to format the response
        embed = discord.Embed(
            title=f"Top {number} Items by {criterion.capitalize()}",
            description=
            f"Here are the top {number} items sorted by {criterion}.",
            color=discord.Color.blue())

        # Add each item as a field in the embed
        for index, (_, row) in enumerate(sorted_data.iterrows(), start=1):
            embed.add_field(
                name=f"{index}. {row['Item name']}",
                value=f"**{criterion.capitalize()}**: {row[column_name]}",
                inline=False)

        # Send the embed as a reply to the user
        await ctx.reply(embed=embed, mention_author=False)

    except Exception as e:
        await ctx.reply(
            "An error occurred while fetching the top items. Please try again later.",
            mention_author=False)
        print(e)


# Command: Filter items based on a given condition (supports Python-like syntax)
@bot.command(name="filter")
async def filter_items(ctx, *, condition: str):
    try:
        # Map shorthand names to actual column names
        column_aliases = {
            "demand": "Demand (out of 10)",
            "value": "Value",
            "rate_of_change": "rate of change"
        }

        # Replace shorthand with actual column names
        for alias, actual_name in column_aliases.items():
            condition = condition.replace(alias, f"`{actual_name}`")

        # Ensure columns are numeric
        for column in ["Demand (out of 10)", "Value", "rate of change"]:
            if column in data.columns:
                data[column] = pd.to_numeric(data[column], errors='coerce')

        # Apply the query
        filtered_data = data.query(condition)

        if filtered_data.empty:
            await ctx.reply("No items match the given filter criteria.",
                            mention_author=False)
            return

        # Format and send results
        embed = discord.Embed(
            title="Filtered Items",
            description=f"Items matching the filter: `{condition}`",
            color=discord.Color.green())

        for _, row in filtered_data.head(10).iterrows():
            embed.add_field(
                name=row["Item name"],
                value=
                f"**Demand**: {row['Demand (out of 10)']}\n**Value**: {row['Value']}",
                inline=False)

        if len(filtered_data) > 10:
            embed.set_footer(
                text=
                "Showing the first 10 results. Refine your filter for more specific results."
            )

        await ctx.reply(embed=embed, mention_author=False)

    except Exception as e:
        await ctx.reply(
            "Invalid filter criteria. Use Python-like syntax, e.g., `demand > 8`.",
            mention_author=False)
        print(f"Error: {e}")


# Command: Compare two items side by side (supports emojis as input)
# Command: Compare two items side by side (supports emojis as input)
@bot.command(name="compare")
async def compare(ctx, *, trade_details: str = None):
    """
    Compare trades with simplified input format.
    Format: `!compare [my_items] <:for:emoji_id> [their_items]`
    Example: `!compare <:FrostSSJ4_Aura:1310698865510584411> x2, <:FestivalAura:1310704728765632644> x1 <:for:1310746627572633664> <:FestivalAura:1310704728765632644> x3`
    """
    if not trade_details:
        await ctx.send("Please provide trade details in this format:\n"
                       "`!compare [my_items] <:for:emoji_id> [their_items]`")
        return

    try:
        # Split the input into "my trade" and "their trade" using the specific emoji as a separator
        if "<:for:1310746627572633664>" not in trade_details:
            await ctx.send(
                "Invalid format! Use `!compare [my_items] <:for:emoji_id> [their_items]`"
            )
            return

        my_trade_str, their_trade_str = trade_details.split(
            "<:for:1310746627572633664>")

        # Convert each trade string into item-quantity pairs
        def parse_items(trade_str):
            items = trade_str.split(",")
            parsed = []
            for item in items:
                parts = item.strip().split("x")
                if len(parts) == 2:
                    parsed.append((parts[0].strip(), int(parts[1].strip())))
            return parsed

        my_trade = parse_items(my_trade_str)
        their_trade = parse_items(their_trade_str)

        # Calculate trade values
        def calculate_trade_value(trade):
            total_value = 0
            details = []
            for item_name, quantity in trade:
                item_data = find_exact_or_closest(item_name)
                if item_data is not None:
                    item_value = int(item_data['Value'])
                    total_value += item_value * quantity
                    details.append(
                        f"{item_name} x{quantity} - Value: {item_value:>5} each (Total: {item_value * quantity:>6})"
                    )
                else:
                    details.append(f"{item_name} x{quantity} - Not Found")
            return total_value, details

        my_trade_value, my_trade_details = calculate_trade_value(my_trade)
        their_trade_value, their_trade_details = calculate_trade_value(
            their_trade)

        # Calculate fairness
        fairness = "Fair Trade!!!" if my_trade_value == their_trade_value else (
            "Your Trade is Overpaying!" if my_trade_value > their_trade_value
            else "Their Trade is Overpaying!")

        # Create embed
        embed = discord.Embed(title="Trade Comparison",
                              color=discord.Color.blue())
        embed.add_field(name="Your Trade",
                        value="\n".join(my_trade_details) +
                        f"\n**Total Value**: {my_trade_value:>6}",
                        inline=False)
        embed.add_field(name="Their Trade",
                        value="\n".join(their_trade_details) +
                        f"\n**Total Value**: {their_trade_value:>6}",
                        inline=False)
        embed.add_field(name="Comparison Result", value=fairness, inline=False)

        await ctx.reply(embed=embed)

    except Exception as e:
        await ctx.send(
            "An error occurred while processing the trade. Please check your input format."
        )
        print(e)


# Command: List recent updates based on the rate of change
@bot.command(name="recent")
async def recent_items(ctx, number: int = 5):
    try:
        if number <= 0:
            await ctx.reply(
                "Please provide a positive number of items to display.",
                mention_author=False)
            return

        recent_data = data.tail(number)

        embed = discord.Embed(title="Recent Items",
                              description=f"Showing the last {number} items.",
                              color=discord.Color.blue())

        for _, row in recent_data.iterrows():
            embed.add_field(
                name=row["Item name"],
                value=
                f"**Demand**: {row['Demand (out of 10)']}\n**Value**: {row['Value']}",
                inline=False)

        await ctx.reply(embed=embed, mention_author=False)

    except Exception as e:
        await ctx.reply(
            "An error occurred while fetching recent items. Please try again.",
            mention_author=False)
        print(f"Error: {e}")


# Custom Help Command: Displays the list of available bot commands
bot.remove_command(
    "help")  # Remove the default help command to replace with a custom one


@bot.command(name="help")
async def custom_help(ctx):
    # Create an embed to format the help text
    embed = discord.Embed(
        title="Watashi Help Menu",
        description="Here are the available commands and how to use them:",
        color=discord.Color.green(),
    )

    # Add fields for each command category
    embed.add_field(
        name="🔍 Value Commands",
        value=
        "`!value [item name]` or `!v [item name]`: Get the value, demand, and rate of change for an item.",
        inline=False)
    embed.add_field(
        name="📊 Top Items",
        value=
        "`!top [number] [criterion]`: List the top items based on `demand` or `value`.\nExample: `!top 5 demand`.",
        inline=False)
    embed.add_field(
        name="🔎 Filter Items",
        value=
        "`!filter [condition]`: Find items matching specific criteria.\nExample: `!filter demand > 8`.",
        inline=False)
    embed.add_field(
        name="⚔ Compare Items",
        value="`!compare [item1] [item2]`: Compare two items side by side.",
        inline=False)
    embed.add_field(
        name="📈 Recent Updates",
        value="`!recent [number]`: List items with the highest rate of change.",
        inline=False)
    embed.add_field(name="❓ Help Command",
                    value="`!help`: Show this help message.",
                    inline=False)

    # Add a footer
    embed.set_footer(text="Use these commands to get insights about items! 🚀")

    # Reply to the user's message with the embed
    await ctx.reply(embed=embed, mention_author=False)


# Command: Get value, demand, and rate of change for a specific item
@bot.command(name="value", aliases=["v"])
async def value(ctx, *, item_name: str = None):
    # Ensure the user provides an item name
    if not item_name:
        await ctx.reply(
            "Please provide an item name. Example: `!value Frost Aura`",
            mention_author=False)
        return

    # Map emojis to item names if applicable
    emoji_map = {
        "<:FrostSSJ4_Aura:1310698865510584411>":
        "Frost aura",  # Example mapping
        "<:FestivalAura:1310704728765632644>": "Festival aura",
    }

    # Convert emoji input to corresponding item name
    if item_name in emoji_map:
        item_name = emoji_map[item_name]

    # Search for the item in the data
    result = enhanced_find_item(data, item_name, case_sensitive=False)

    if "Item not found" in result:  # If the item is not found
        await ctx.reply(result, mention_author=False)
        return

    # Parse the item's details
    exact_match = find_exact_or_closest(item_name)
    if exact_match is None:
        await ctx.reply("Item not found. Please check the name and try again.",
                        mention_author=False)
        return

    # Create a Discord embed with the item details
    embed = discord.Embed(
        title="Item Details",
        description=f"Information for **{exact_match['Item name']}**",
        color=discord.Color.blue(),
    )
    embed.add_field(name="Demand",
                    value=f"{exact_match['Demand (out of 10)']}/10",
                    inline=True)
    embed.add_field(name="Value", value=f"{exact_match['Value']}", inline=True)
    embed.add_field(name="Rate of Change",
                    value=f"{exact_match['rate of change']}",
                    inline=False)

    # Add footer or any additional notes
    embed.set_footer(text="Use the !compare command to compare items!")

    # Reply to the triggering message
    await ctx.reply(embed=embed, mention_author=False)


# Event: When the bot is ready, log the event
@bot.event
async def on_ready():
    print(f"Logged in as {bot.user}")


# Event: Capture and process messages
@bot.event
async def on_message(message):
    if message.author.bot:
        return  # Ignore messages from other bots
    print(
        f"[{message.author}] {message.content}")  # Log messages for debugging
    await bot.process_commands(message)


# Run the bot with the provided token (replace 'token' with the actual bot token)

# Run the bot with the provided token (replace 'token' with the actual bot token)
bot.run(os.getenv('TOKEN'))  # Replace with your actual bot token
