
import discord
from discord.ext import commands
from discord import app_commands
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import json
import random
import asyncio
import re
import time
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def find_user_in_sheets(user_id, sheets=["All Signups", "All Winners"]):
    for sheet_name in sheets:
        sheet = get_sheet(sheet_name)
        if not sheet:
            continue
        records = get_all_records_safe(sheet)
        for i, record in enumerate(records, start=2):
            if str(record.get("User ID", "")) == str(user_id):
                return sheet, i, record
    return None, None, None

# Initialize Discord bot
intents = discord.Intents.default()
intents.message_content = True

class ElectionBot(commands.Bot):
    def __init__(self):
        super().__init__(command_prefix='/', intents=intents)

    async def setup_hook(self):
        # Sync commands with Discord
        try:
            synced = await self.tree.sync()
            logger.info(f"Synced {len(synced)} command(s)")
        except Exception as e:
            logger.error(f"Failed to sync commands: {e}")

bot = ElectionBot()

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
try:
    creds_json = os.getenv('GOOGLE_CREDS')
    if not creds_json:
        raise ValueError("GOOGLE_CREDS environment variable not set")
    creds_dict = json.loads(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet_id = "16CwRg1p2w0kU0xLHZbN1CnxDgy0JhKy8KbBZSxqIfHo"
    spreadsheet = client.open_by_key(spreadsheet_id)
except gspread.exceptions.SpreadsheetNotFound:
    logger.error("Spreadsheet not found. Check the spreadsheet ID and permissions.")
    exit(1)
except json.JSONDecodeError:
    logger.error("Invalid GOOGLE_CREDS format. Ensure it's valid JSON.")
    exit(1)
except Exception as e:
    logger.error(f"Error initializing Google Sheets: {e}")
    exit(1)

# Helper functions
def get_sheet(sheet_name, create_if_missing=False, cols=10):
    try:
        return spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        logger.warning(f"Worksheet '{sheet_name}' not found.")
        if create_if_missing:
            return spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=cols)
        return None

def retry_api_call(func, *args, max_retries=3, initial_delay=1, **kwargs):
    for attempt in range(max_retries):
        try:
            return func(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            error_msg = str(e)
            if "invalid authentication" in error_msg.lower():
                logger.error("Google Sheets authentication failed. Check GOOGLE_CREDS.")
                raise ValueError("Authentication error: Invalid GOOGLE_CREDS or service account permissions.")
            if attempt == max_retries - 1:
                logger.error(f"API call failed after {max_retries} attempts: {e}")
                raise
            delay = initial_delay * (2 ** attempt)
            logger.warning(f"APIError on attempt {attempt + 1}, retrying after {delay}s: {e}")
            time.sleep(delay)

def get_current_cycle():
    cycles_sheet = get_sheet("Cycles")
    if not cycles_sheet:
        return None
    try:
        return {
            "year": int(retry_api_call(cycles_sheet.acell, "G2").value or 1990),
            "signups_open": retry_api_call(cycles_sheet.acell, "G3").value.lower() == "true",
            "phase": retry_api_call(cycles_sheet.acell, "G4").value or "Primary"
        }
    except Exception as e:
        logger.error(f"Error getting current cycle: {e}")
        return None

def update_cycle_year(new_year):
    cycles_sheet = get_sheet("Cycles")
    if cycles_sheet:
        retry_api_call(cycles_sheet.update_acell, "G2", new_year)

def update_signups_status(status):
    cycles_sheet = get_sheet("Cycles")
    if cycles_sheet:
        retry_api_call(cycles_sheet.update_acell, "G3", str(status).upper())

def update_phase(phase):
    cycles_sheet = get_sheet("Cycles")
    if cycles_sheet:
        retry_api_call(cycles_sheet.update_acell, "G4", phase)

def get_eligible_seats(year, state=None):
    cycles_sheet = get_sheet("Cycles")
    if not cycles_sheet:
        return []
    records = get_all_records_safe(cycles_sheet)
    eligible_seats = [r for r in records if r.get("Year", 0) == year]
    if state:
        logger.info(f"Querying seats for state: {state}, year: {year}")
        eligible_seats = [r for r in eligible_seats if r.get("State", "").lower() == state.lower()]
    logger.info(f"Found {len(eligible_seats)} eligible seats: {[r.get('Seat ID', '') for r in eligible_seats]}")
    return eligible_seats

def add_signup(user_id, name, seat_id, party, state, office):
    signups_sheet = get_sheet("All Signups")
    if not signups_sheet:
        return
    cycle = get_current_cycle()
    if not cycle:
        return
    row = [
        user_id,
        name,
        seat_id,
        party,
        cycle["phase"],
        office,
        0,  # Corruption
        100,  # Stamina
        0,  # Points
        ""  # Winner
    ]
    retry_api_call(signups_sheet.append_row, row)

def update_term_year(seat_id, current_year, term):
    cycles_sheet = get_sheet("Cycles")
    if not cycles_sheet:
        return
    records = get_all_records_safe(cycles_sheet)
    for i, record in enumerate(records, start=2):
        if record.get("Seat ID", "") == seat_id:
            term = int(record.get("Term/Year", 0)) if str(record.get("Term/Year", 0)).isdigit() else 0
            new_year = current_year + term
            retry_api_call(cycles_sheet.update_cell, i, 4, new_year)
            break

def get_all_records_safe(sheet, expected_headers=None):
    try:
        # First check if the sheet has any data
        all_values = retry_api_call(sheet.get_all_values)
        if not all_values or len(all_values) < 2:
            return []
        
        # Clean the header row to remove empty strings and duplicates
        header_row = all_values[0]
        cleaned_headers = []
        for i, header in enumerate(header_row):
            if header.strip():  # Only add non-empty headers
                cleaned_headers.append(header.strip())
            else:
                # For empty headers, create a placeholder name
                cleaned_headers.append(f"Column_{i+1}")
        
        # Remove duplicates by adding numbers to duplicate names
        final_headers = []
        header_count = {}
        for header in cleaned_headers:
            if header in header_count:
                header_count[header] += 1
                final_headers.append(f"{header}_{header_count[header]}")
            else:
                header_count[header] = 0
                final_headers.append(header)
        
        # Update the sheet with cleaned headers if needed
        if final_headers != header_row:
            retry_api_call(sheet.update, 'A1', [final_headers])
        
        # Use the cleaned headers to get records
        return retry_api_call(sheet.get_all_records, expected_headers=final_headers)
    except Exception as e:
        logger.error(f"Error getting all records: {e}")
        return []

# Bot events
@bot.event
async def on_ready():
    logger.info(f'Logged in as {bot.user}')

# Slash Commands
@bot.tree.command(name="signup", description="Sign up for an election")
@app_commands.describe(
    state="The state you want to run in",
    name="Your candidate's name", 
    party="Your political party"
)
@app_commands.choices(party=[
    app_commands.Choice(name="Democrats", value="Democrats"),
    app_commands.Choice(name="Republicans", value="Republicans"),
    app_commands.Choice(name="Independent", value="Independent")
])
@app_commands.choices(state=[
    app_commands.Choice(name="Columbia", value="Columbia"),
    app_commands.Choice(name="Cambridge", value="Cambridge"),
    app_commands.Choice(name="Austin", value="Austin"),
    app_commands.Choice(name="Superior", value="Superior"),
    app_commands.Choice(name="Heartland", value="Heartland"),
    app_commands.Choice(name="Yellowstone", value="Yellowstone"),
    app_commands.Choice(name="Phoenix", value="Phoenix")
])
async def signup(interaction: discord.Interaction, state: str, name: str, party: str):
    try:
        await interaction.response.defer()

        signups_sheet = get_sheet("All Signups")
        winners_sheet = get_sheet("All Winners")
        if not signups_sheet:
            await interaction.followup.send("Error accessing signups data.")
            return

        user_id_str = str(interaction.user.id)

        # Check if user is already signed up in All Signups sheet
        signups_records = get_all_records_safe(signups_sheet)
        for record in signups_records:
            record_user_id = str(record.get("User ID", ""))
            if record_user_id == user_id_str and record.get("Winner", "") not in ["Withdrawn", "Loser"]:
                await interaction.followup.send("You are already signed up! Use `/withdraw` to cancel your current candidacy.")
                return

        # Check if user is already in All Winners sheet
        if winners_sheet:
            winners_records = get_all_records_safe(winners_sheet)
            for record in winners_records:
                record_user_id = str(record.get("User ID", ""))
                if record_user_id == user_id_str:
                    await interaction.followup.send("You already have an active candidacy in the current election cycle!")
                    return

        cycle = get_current_cycle()
        if not cycle:
            await interaction.followup.send("Error accessing cycle data.")
            return
        if not cycle["signups_open"]:
            await interaction.followup.send("Signups are closed!")
            return

        seats = get_eligible_seats(cycle["year"], state)
        if not seats:
            await interaction.followup.send(f"No eligible seats for {state} in {cycle['year']}.")
            return

        # Create seat selection view
        class SeatView(discord.ui.View):
            def __init__(self, seats_list):
                super().__init__(timeout=60)
                self.seats = seats_list
                self.add_seat_buttons()

            def add_seat_buttons(self):
                for seat in self.seats[:25]:  # Discord limits to 25 buttons
                    button = discord.ui.Button(
                        label=f"{seat['Seat ID']} ({seat['Office']})",
                        custom_id=seat['Seat ID']
                    )
                    button.callback = self.seat_callback
                    self.add_item(button)

            async def seat_callback(self, button_interaction):
                try:
                    # Acknowledge the interaction immediately
                    await button_interaction.response.defer()

                    seat_id = button_interaction.data['custom_id']
                    selected_seat = next((s for s in self.seats if s.get("Seat ID", "") == seat_id), None)

                    if not selected_seat:
                        await button_interaction.followup.send("Invalid seat selection!", ephemeral=True)
                        return

                    # Triple-check for duplicate signup across both sheets
                    signups_records = get_all_records_safe(signups_sheet)
                    winners_sheet = get_sheet("All Winners")
                    winners_records = get_all_records_safe(winners_sheet) if winners_sheet else []

                    # Check All Signups sheet
                    for record in signups_records:
                        record_user_id = str(record.get("User ID", ""))
                        if record_user_id == user_id_str and record.get('Winner', '') not in ['Withdrawn', 'Loser']:
                            await button_interaction.followup.send("You are already signed up for another race!", ephemeral=True)
                            return

                    # Check All Winners sheet
                    for record in winners_records:
                        record_user_id = str(record.get("User ID", ""))
                        if record_user_id == user_id_str:
                            await button_interaction.followup.send("You already have an active candidacy in the Winners sheet!", ephemeral=True)
                            return

                    add_signup(interaction.user.id, name, seat_id, party, state, selected_seat["Office"])
                    await button_interaction.followup.send(
                        f"{name} signed up for {seat_id} ({selected_seat['Office']}) in {state} as {party}!"
                    )
                except Exception as e:
                    logger.error(f"Error in seat_callback: {e}")
                    try:
                        await button_interaction.followup.send("An error occurred during seat selection. Please try again.", ephemeral=True)
                    except:
                        pass

        seat_options = [f"{s['Seat ID']} ({s['Office']})" for s in seats]
        view = SeatView(seats)
        await interaction.followup.send(
            f"Available seats in {state}:\n" + "\n".join(seat_options) + 
            "\nPlease click the button for the seat you want to run for:",
            view=view
        )

    except Exception as e:
        logger.error(f"Error in signup command: {e}")
        await interaction.followup.send("An error occurred during signup. Please try again or contact an admin.")

@bot.tree.command(name="withdraw", description="Withdraw from an election")
@app_commands.describe(character_name="The name of the character to withdraw")
async def withdraw(interaction: discord.Interaction, character_name: str):
    try:
        await interaction.response.defer()

        signups_sheet = get_sheet("All Signups")
        if not signups_sheet:
            await interaction.followup.send("Error accessing signups data.")
            return

        records = get_all_records_safe(signups_sheet)
        user_id_str = str(interaction.user.id)
        found = False

        for i, record in enumerate(records, start=2):
            record_name = record.get("Name", "").strip()
            record_user_id = str(record.get("User ID", ""))
            if record_name.lower() == character_name.strip().lower() and record_user_id == user_id_str:
                if record.get("Winner", "") in ["Winner", "Loser"]:
                    await interaction.followup.send(f"Cannot withdraw {character_name}. Candidacy is already marked as {record.get('Winner', '')}.")
                    return
                retry_api_call(signups_sheet.delete_rows, i)
                found = True
                await interaction.followup.send(f"{character_name} has withdrawn their candidacy and their record has been removed.")
                break

        if not found:
            await interaction.followup.send(f"No signup found for {character_name} associated with your account.")

    except Exception as e:
        logger.error(f"Error in withdraw command: {e}")
        await interaction.followup.send("An error occurred during withdrawal. Please try again or contact an admin.")

@bot.tree.command(name="rally", description="Hold a rally to gain points")
async def rally(interaction: discord.Interaction):
    try:
        await interaction.response.defer()

        cycle = get_current_cycle()
        if not cycle:
            await interaction.followup.send("Error accessing cycle data.")
            return

        sheet = get_sheet("All Signups") if cycle["phase"] == "Primary" else get_sheet("All Winners")
        if not sheet:
            await interaction.followup.send("Error accessing data sheet.")
            return

        records = get_all_records_safe(sheet)
        for i, record in enumerate(records, start=2):
            if str(record.get("User ID", "")) == str(interaction.user.id) and record.get("Winner", "") != "Withdrawn":
                cost = 10
                if record.get("Stamina", 0) < cost:
                    await interaction.followup.send(f"Not enough stamina! Need {cost}, have {max(record.get('Stamina', 0), 0)}.")
                    return

                points = random.randint(5, 15)
                stamina = max(record.get("Stamina", 0) - cost, 0)
                new_points = record.get("Points", 0) + points
                retry_api_call(sheet.update_cell, i, 8, stamina)
                retry_api_call(sheet.update_cell, i, 9, new_points)
                await interaction.followup.send(f"{record.get('Name', '')} rallied and gained {points} points! Stamina: {stamina}")
                return

        await interaction.followup.send("You are not signed up or have withdrawn.")

    except Exception as e:
        logger.error(f"Error in rally command: {e}")
        await interaction.followup.send("An error occurred during rally. Please try again or contact an admin.")

@bot.tree.command(name="canvassing", description="Go door-to-door canvassing for votes")
async def canvassing(interaction: discord.Interaction):
    try:
        await interaction.response.defer()

        cycle = get_current_cycle()
        if not cycle:
            await interaction.followup.send("Error accessing cycle data.")
            return

        sheet = get_sheet("All Signups") if cycle["phase"] == "Primary" else get_sheet("All Winners")
        if not sheet:
            await interaction.followup.send("Error accessing data sheet.")
            return

        records = get_all_records_safe(sheet)
        for i, record in enumerate(records, start=2):
            if str(record.get("User ID", "")) == str(interaction.user.id) and record.get("Winner", "") != "Withdrawn":
                cost = 5
                if record.get("Stamina", 0) < cost:
                    await interaction.followup.send(f"Not enough stamina! Need {cost}, have {max(record.get('Stamina', 0), 0)}.")
                    return

                points = random.randint(3, 10)
                stamina = max(record.get("Stamina", 0) - cost, 0)
                new_points = record.get("Points", 0) + points
                retry_api_call(sheet.update_cell, i, 8, stamina)
                retry_api_call(sheet.update_cell, i, 9, new_points)
                await interaction.followup.send(f"{record.get('Name', '')} went canvassing and gained {points} points! Stamina: {stamina}")
                return

        await interaction.followup.send("You are not signed up or have withdrawn.")

    except Exception as e:
        logger.error(f"Error in canvassing command: {e}")
        await interaction.followup.send("An error occurred during canvassing. Please try again or contact an admin.")

@bot.tree.command(name="poll", description="Check polling results")
@app_commands.describe(scope="Check national polls or specify a state")
async def poll(interaction: discord.Interaction, scope: str = "nation"):
    try:
        await interaction.response.defer()

        cycle = get_current_cycle()
        if not cycle:
            await interaction.followup.send("Error accessing cycle data.")
            return

        sheet = get_sheet("All Signups") if cycle["phase"] == "Primary" else get_sheet("All Winners")
        if not sheet:
            await interaction.followup.send("Error accessing data sheet.")
            return

        records = get_all_records_safe(sheet)

        if scope.lower() == "nation":
            candidates = [r for r in records if r.get("Phase", "") == cycle["phase"] and r.get("Winner", "") != "Withdrawn"]
            if not candidates:
                await interaction.followup.send("No candidates to poll.")
                return
            total_points = sum(c.get("Points", 0) for c in candidates)
            if total_points == 0:
                await interaction.followup.send("No polling data available.")
                return
            message = "National Poll Results:\n"
            for c in candidates:
                percentage = (c.get("Points", 0) / total_points) * 100
                message += f"{c.get('Name', '')} ({c.get('Party', '')}): {percentage:.1f}%\n"
            await interaction.followup.send(message)
        else:
            candidates = [r for r in records if r.get("State", "").lower() == scope.lower() and r.get("Phase", "") == cycle["phase"] and r.get("Winner", "") != "Withdrawn"]
            if not candidates:
                await interaction.followup.send(f"No candidates in {scope} to poll.")
                return
            total_points = sum(c.get("Points", 0) for c in candidates)
            if total_points == 0:
                await interaction.followup.send(f"No polling data available in {scope}.")
                return
            message = f"Poll Results in {scope}:\n"
            for c in candidates:
                percentage = (c.get("Points", 0) / total_points) * 100
                message += f"{c.get('Name', '')} ({c.get('Party', '')}): {percentage:.1f}%\n"
            await interaction.followup.send(message)

    except Exception as e:
        logger.error(f"Error in poll command: {e}")
        await interaction.followup.send("An error occurred while polling. Please try again or contact an admin.")

@bot.tree.command(name="commands", description="Show all available commands organized by category")
async def commands(interaction: discord.Interaction):
    try:
        await interaction.response.defer()

        embed = discord.Embed(title="🗳️ APRP Election Bot Commands", color=0x0099ff)

        # Player Commands
        player_commands = [
            "`/signup` - Sign up for an election",
            "`/withdraw` - Withdraw from an election",
            "`/rally` - Hold a rally to gain points (costs 10 stamina)",
            "`/canvassing` - Go door-to-door canvassing (costs 5 stamina)",
            "`/poll` - Check polling results (national or by state)",
            "`/edit_signup` - Edit your party affiliation"
        ]
        embed.add_field(name="📊 Player Commands", value="\n".join(player_commands), inline=False)

        # Campaign Commands
        campaign_commands = [
            "`/ad` - Run a campaign ad (costs 15 stamina)",
            "`/poster` - Post campaign posters (costs 10 stamina)",
            "`/hall` - Hold a town hall (costs 20 stamina)",
            "`/platform` - Set your presidential platform (costs 50 stamina)",
            "`/endorse` - Endorse another candidate",
            "`/donor` - Gain points from donors (increases corruption, costs 15 stamina)",
            "`/special` - Give a special interest speech (increases corruption, costs 20 stamina)"
        ]
        embed.add_field(name="🎪 Campaign Commands", value="\n".join(campaign_commands), inline=False)

        # Admin Commands
        admin_commands = [
            "`/run_primaries` - Run primary or general elections",
            "`/tally_winners` - Tally up points and determine winners by seat/party",
            "`/transfer_winners` - Transfer declared winners to All Winners sheet",
            "`/open_signups` - Open signups for elections",
            "`/close_signups` - Close signups for elections",
            "`/advance_cycle` - Advance to the next election year",
            "`/list_signups` - List all current signups",
            "`/generate_ballot` - Generate a ballot for a given year",
            "`/archive_signups` - Archive completed signups",
            "`/debug_candidate` - Debug a specific candidate"
        ]
        embed.add_field(name="🔧 Admin Commands", value="\n".join(admin_commands), inline=False)

        embed.add_field(name="📋 General Info", value="`/commands` - Show this help menu", inline=False)

        embed.set_footer(text="💡 Tip: Stamina regenerates over time. Manage your corruption levels!")

        await interaction.followup.send(embed=embed)

    except Exception as e:
        logger.error(f"Error in commands command: {e}")
        await interaction.followup.send("An error occurred while displaying commands. Please try again or contact an admin.")

# Admin Commands
@bot.tree.command(name="tally_winners", description="[ADMIN] Tally up points and determine winners by seat and party")
async def tally_winners(interaction: discord.Interaction):
    try:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command!", ephemeral=True)
            return

        await interaction.response.defer()

        signups_sheet = get_sheet("All Signups")
        if not signups_sheet:
            await interaction.followup.send("Error accessing signups data.")
            return

        records = get_all_records_safe(signups_sheet)

        # Group candidates by seat ID and party
        races = {}
        for record in records:
            if record.get("Winner", "") in ["Withdrawn", "Winner", "Loser"]:
                continue

            seat_id = record.get("Seat ID", "")
            party = record.get("Party", "")
            key = (seat_id, party)

            if key not in races:
                races[key] = []
            races[key].append(record)

        winners_declared = 0

        # Determine winners within each party for each seat
        for (seat_id, party), candidates in races.items():
            if not candidates:
                continue

            # Find the candidate with the highest points
            winner = max(candidates, key=lambda x: x.get("Points", 0))

            # Update the records in the sheet
            for i, record in enumerate(records, start=2):
                if (record.get("Seat ID", "") == seat_id and 
                    record.get("Party", "") == party and
                    record.get("Winner", "") not in ["Withdrawn", "Winner", "Loser"]):

                    if record.get("Name", "") == winner.get("Name", ""):
                        retry_api_call(signups_sheet.update_cell, i, 10, "Winner")
                        winners_declared += 1
                    else:
                        retry_api_call(signups_sheet.update_cell, i, 10, "Loser")

        await interaction.followup.send(f"Tallying complete! {winners_declared} winners declared across all seats and parties.")

    except Exception as e:
        logger.error(f"Error in tally_winners command: {e}")
        await interaction.followup.send("An error occurred while tallying winners. Please try again or contact an admin.")

@bot.tree.command(name="transfer_winners", description="[ADMIN] Transfer declared winners from All Signups to All Winners sheet")
async def transfer_winners(interaction: discord.Interaction):
    try:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command!", ephemeral=True)
            return

        await interaction.response.defer()

        signups_sheet = get_sheet("All Signups")
        winners_sheet = get_sheet("All Winners")

        if not signups_sheet or not winners_sheet:
            await interaction.followup.send("Error accessing required sheets.")
            return

        cycle = get_current_cycle()
        if not cycle:
            await interaction.followup.send("Error accessing cycle data.")
            return

        records = get_all_records_safe(signups_sheet)
        transferred = 0

        # Check if winners sheet has headers, add them if missing
        winners_records = get_all_records_safe(winners_sheet)
        if not winners_records:
            headers = ["Year", "Office", "State", "District", "Candidate", "Party", "Points", "Votes", "Corruption", "Final Score", "Winner"]
            retry_api_call(winners_sheet.append_row, headers)

        for record in records:
            if record.get("Winner", "") == "Winner":
                # Check if this winner is already in the All Winners sheet
                user_id = record.get("User ID", "")
                seat_id = record.get("Seat ID", "")

                already_exists = False
                for winner_record in winners_records:
                    if (str(winner_record.get("User ID", "")) == str(user_id) and 
                        winner_record.get("Seat ID", "") == seat_id):
                        already_exists = True
                        break

                if not already_exists:
                    # Parse district from seat ID (e.g., REP-CO-1 -> District 1)
                    district = ""
                    if "-" in seat_id:
                        parts = seat_id.split("-")
                        if len(parts) >= 3:
                            district = f"District {parts[-1]}"

                    # Transfer to All Winners sheet
                    winner_row = [
                        cycle["year"],
                        record.get("Office", ""),
                        record.get("State", ""),
                        district,
                        record.get("Name", ""),
                        record.get("Party", ""),
                        record.get("Points", 0),
                        0,  # Votes (placeholder)
                        record.get("Corruption", 0),
                        record.get("Points", 0),  # Final Score = Points for now
                        "Yes"
                    ]

                    retry_api_call(winners_sheet.append_row, winner_row)
                    transferred += 1

        await interaction.followup.send(f"Transfer complete! {transferred} winners moved to All Winners sheet.")

    except Exception as e:
        logger.error(f"Error in transfer_winners command: {e}")
        await interaction.followup.send("An error occurred while transferring winners. Please try again or contact an admin.")

@bot.tree.command(name="open_signups", description="[ADMIN] Open signups for elections")
async def open_signups(interaction: discord.Interaction):
    try:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command!", ephemeral=True)
            return

        await interaction.response.defer()
        
        update_signups_status(True)
        await interaction.followup.send("Election signups are now OPEN! 🗳️")

    except Exception as e:
        logger.error(f"Error in open_signups command: {e}")
        await interaction.followup.send("An error occurred while opening signups. Please try again or contact an admin.")

@bot.tree.command(name="close_signups", description="[ADMIN] Close signups for elections")
async def close_signups(interaction: discord.Interaction):
    try:
        if not interaction.user.guild_permissions.administrator:
            await interaction.response.send_message("You need administrator permissions to use this command!", ephemeral=True)
            return

        await interaction.response.defer()
        
        update_signups_status(False)
        await interaction.followup.send("Election signups are now CLOSED! 🚫")

    except Exception as e:
        logger.error(f"Error in close_signups command: {e}")
        await interaction.followup.send("An error occurred while closing signups. Please try again or contact an admin.")

# Run the bot
if __name__ == "__main__":
    bot_token = os.getenv('DISCORD_TOKEN')
    if not bot_token:
        logger.error("DISCORD_TOKEN environment variable not set")
        exit(1)
    bot.run(bot_token)
