const { Client, GatewayIntentBits, SlashCommandBuilder, EmbedBuilder, ActionRowBuilder, ButtonBuilder, ButtonStyle } = require('discord.js');
const { google } = require('googleapis');

// Initialize storage
const pollVotes = new Map();
const pollData = new Map();

// Initialize Discord client
const client = new Client({
    intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent
    ]
});

// Google Sheets setup
let sheets;
const SPREADSHEET_ID = "16CwRg1p2w0kU0xLHZbN1CnxDgy0JhKy8KbBZSxqIfHo";

async function initializeGoogleSheets() {
    try {
        const credsJson = process.env.GOOGLE_CREDS;
        if (!credsJson) {
            throw new Error("GOOGLE_CREDS environment variable not set");
        }

        const creds = JSON.parse(credsJson);
        const auth = new google.auth.GoogleAuth({
            credentials: creds,
            scopes: [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        });

        sheets = google.sheets({ version: 'v4', auth });
        console.log('Google Sheets initialized successfully');
    } catch (error) {
        console.error('Error initializing Google Sheets:', error);
        process.exit(1);
    }
}

// Helper functions
async function getSheetData(sheetName) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: normalizedSheetName,
        });
        return response.data.values || [];
    } catch (error) {
        if (error.code === 404) {
            console.error(`Sheet "${sheetName}" not found. Creating it...`);
            await createSheetIfNotExists(sheetName);
            return [];
        }
        console.error(`Error getting sheet data for ${sheetName}:`, error);
        return [];
    }
}

function normalizeSheetName(sheetName) {
    if (!sheetName) return '';
    return sheetName.replace(/\b\w/g, c => c.toUpperCase());
}

async function createSheetIfNotExists(sheetName) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        const spreadsheet = await sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID
        });

        const sheetExists = spreadsheet.data.sheets.some(
            sheet => sheet.properties.title === normalizedSheetName
        );

        if (!sheetExists) {
            console.log(`Creating sheet: ${normalizedSheetName}`);
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: SPREADSHEET_ID,
                resource: {
                    requests: [{
                        addSheet: {
                            properties: {
                                title: normalizedSheetName
                            }
                        }
                    }]
                }
            });

            await initializeSheetHeaders(normalizedSheetName);
        }
    } catch (error) {
        console.error(`Error creating sheet ${sheetName}:`, error);
    }
}

async function initializeSheetHeaders(sheetName) {
    try {
        let headers = [];

        switch (sheetName.toLowerCase()) {
            case 'all signups':
                headers = ["User ID", "Name", "Seat ID", "Party", "Phase", "States", "Office", "Corruption", "Stamina", "Points", "Winner"];
                break;
            case 'all winners':
                headers = ["Year", "Office", "State", "Seat ID", "Candidate", "Party", "Points", "Votes", "Corruption", "Final Score", "Winner"];
                break;
            case 'cycles':
                headers = [
                    ["Seat ID", "Office", "State", "Year", "Term/Year", "SETTING", ""],
                    ["SEN-CO-3", "Senate", "Columbia", "1990", "6", "YEAR", ""],
                    ["", "", "", "", "", "CYCLE", ""],
                    ["", "", "", "", "", "MONTH", ""],
                    ["", "", "", "", "", "PHASE", ""]
                ];
                break;
            default:
                headers = ["User ID", "Name", "Seat ID", "Party", "Phase", "States", "Office", "Corruption", "Stamina", "Points", "Winner"];
        }

        if (sheetName.toLowerCase() === 'cycles') {
            for (const row of headers) {
                await appendToSheet(sheetName, row);
            }
        } else {
            await appendToSheet(sheetName, headers);
        }
    } catch (error) {
        console.error(`Error initializing headers for ${sheetName}:`, error);
    }
}

async function appendToSheet(sheetName, values) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: normalizedSheetName,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [values]
            }
        });
    } catch (error) {
        console.error(`Error appending to sheet ${sheetName}:`, error);
        throw error;
    }
}

async function updateSheetCell(sheetName, row, col, value) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        const range = `${normalizedSheetName}!${String.fromCharCode(65 + col)}${row}`;
        await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: range,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[value]]
            }
        });
    } catch (error) {
        console.error(`Error updating cell ${sheetName}!${range}:`, error);
        throw error;
    }
}

async function deleteSheetRow(sheetName, rowIndex) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        const sheetId = await getSheetId(normalizedSheetName);
        if (sheetId === null) {
            throw new Error(`Sheet ${sheetName} not found`);
        }
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            resource: {
                requests: [{
                    deleteDimension: {
                        range: {
                            sheetId: sheetId,
                            dimension: 'ROWS',
                            startIndex: rowIndex - 1,
                            endIndex: rowIndex
                        }
                    }
                }]
            }
        });
    } catch (error) {
        console.error(`Error deleting row from ${sheetName}:`, error);
        throw error;
    }
}

async function getSheetId(sheetName) {
    try {
        const normalizedSheetName = normalizeSheetName(sheetName);
        const response = await sheets.spreadsheets.get({
            spreadsheetId: SPREADSHEET_ID
        });
        const sheet = response.data.sheets.find(s => s.properties.title === normalizedSheetName);
        return sheet ? sheet.properties.sheetId : null;
    } catch (error) {
        console.error(`Error getting sheet ID for ${sheetName}:`, error);
        return null;
    }
}

function parseSheetData(data) {
    if (!data || data.length === 0) return [];
    const headers = data[0] || [];
    return data.slice(1).map(row => {
        const record = {};
        headers.forEach((header, index) => {
            record[header] = row[index] || '';
        });
        return record;
    });
}

async function getCurrentCycle() {
    try {
        const data = await getSheetData('Cycles');
        if (!data || data.length < 7) { // Header + 6 rows (Cycle info + 5 phases)
            console.log('Creating default cycle settings...');
            return {
                year: 1990,
                cycle: 1,
                month: 8, // Adjusted to August (was July)
                phase: 'Signups',
                isPaused: false
            };
        }

        // Read Cycle Number, Year, Month from F6, G6, H6 (row 6, cols 5, 6, 7)
        let cycle = parseInt(data[5][5]) || 1; // F6
        let year = parseInt(data[5][6]) || 1990; // G6
        let month = parseInt(data[5][7]) || 8; // H6
        let phase = 'Signups';
        let isPaused = data[5][8]?.toLowerCase() === 'true' || false; // I6 for pause

        // Determine phase from phase table (starting at A2)
        const phases = [
            { name: 'Signups', start: 4, end: 8 },        // Adjusted from 3-7 to 4-8
            { name: 'Primary Campaign', start: 8, end: 12 }, // Adjusted from 7-11 to 8-12
            { name: 'Primary Election', start: 12, end: 12 }, // Adjusted from 11-12 to 12-12
            { name: 'General Campaign', start: 4, end: 9 },  // Adjusted from 3-9 to 4-9
            { name: 'General Election', start: 12, end: 12 } // Adjusted from 11-12 to 12-12
        ];

        for (const p of phases) {
            if (month >= p.start && month <= p.end && (year === 1991 || year === 1992)) {
                phase = p.name;
                break;
            }
        }

        return { year, cycle, month, phase, isPaused };
    } catch (error) {
        console.error('Error getting current cycle:', error);
        return {
            year: 1990,
            cycle: 1,
            month: 8,
            phase: 'Signups',
            isPaused: false
        };
    }
}

async function updateCycleYear(newYear) {
    try {
        await updateSheetCell('Cycles', 6, 6, newYear); // G6
    } catch (error) {
        console.error('Error updating cycle year:', error);
    }
}

async function updateCycleMonth(newMonth) {
    try {
        await updateSheetCell('Cycles', 6, 7, newMonth); // H6
    } catch (error) {
        console.error('Error updating cycle month:', error);
    }
}

async function updateCyclePhase(newPhase) {
    try {
        // Update phase indirectly by setting month to match phase range
        const phases = [
            { name: 'Signups', start: 4, end: 8 },
            { name: 'Primary Campaign', start: 8, end: 12 },
            { name: 'Primary Election', start: 12, end: 12 },
            { name: 'General Campaign', start: 4, end: 9 },
            { name: 'General Election', start: 12, end: 12 }
        ];
        const phaseData = phases.find(p => p.name === newPhase);
        if (phaseData) {
            await updateCycleMonth(phaseData.start); // Set to start month of phase
        }
    } catch (error) {
        console.error('Error updating cycle phase:', error);
    }
}

async function updateCycleCycle(newCycle) {
    try {
        await updateSheetCell('Cycles', 6, 5, newCycle); // F6
    } catch (error) {
        console.error('Error updating cycle number:', error);
    }
}

async function updateCyclePause(isPaused) {
    try {
        await updateSheetCell('Cycles', 6, 8, isPaused ? 'TRUE' : 'FALSE'); // I6
    } catch (error) {
        console.error('Error updating cycle pause status:', error);
    }
}

async function getEligibleSeats(year, state = null) {
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Cycles!A1:BZ1000',
        });

        const data = response.data.values || [];
        if (!data || data.length < 2) {
            console.log('No seat data found in Cycles sheet');
            return [];
        }

        let headerRowIndex = -1;
        let headers = [];

        for (let i = 0; i < data.length; i++) {
            if (data[i][0] === 'Seat ID') {
                headerRowIndex = i;
                headers = data[i];
                break;
            }
        }

        if (headerRowIndex === -1) {
            console.log('Could not find header row with "Seat ID"');
            return [];
        }

        const records = [];
        for (let i = headerRowIndex + 1; i < data.length; i++) {
            const row = data[i];
            if (!row || row[0]?.trim() === '') continue;

            const record = {};
            headers.forEach((header, index) => {
                record[header] = row[index] || '';
            });

            if (record['Seat ID'] && record['Seat ID'].trim() !== '') {
                records.push(record);
            }
        }

        console.log(`Found ${records.length} total seats in Cycles sheet`);

        let filteredSeats = records;
        if (state) {
            filteredSeats = records.filter(r => 
                r.State?.toLowerCase() === state.toLowerCase()
            );
            console.log(`Filtered to ${filteredSeats.length} seats in ${state}`);
        }

        let eligibleSeats = filteredSeats.filter(r => {
            const seatYear = parseInt(r.Year) || 0;
            const termLength = parseInt(r['Term/Year']) || 0;
            const currentYear = parseInt(year);

            if (seatYear === 0) {
                console.log(`Including seat ${r['Seat ID']} - no valid year data`);
                return true;
            }

            if (seatYear === currentYear) {
                console.log(`Including seat ${r['Seat ID']} - year matches exactly (${seatYear})`);
                return true;
            }

            if (termLength <= 0) {
                const eligible = seatYear <= currentYear;
                console.log(`Seat ${r['Seat ID']}: no term length, year=${seatYear}, current=${currentYear}, eligible=${eligible}`);
                return eligible;
            }

            const isEligible = (currentYear - seatYear) % termLength === 0;
            console.log(`Seat ${r['Seat ID']}: year=${seatYear}, term=${termLength}, current=${currentYear}, eligible=${isEligible}`);
            return isEligible;
        });

        console.log(`Found ${eligibleSeats.length} eligible seats for year ${year}${state ? ` in ${state}` : ''}`);

        return eligibleSeats;
    } catch (error) {
        console.error('Error getting eligible seats:', error);
        return [];
    }
}

async function findUserInSheets(userId) {
    const sheetsToCheck = ['All Signups', 'All Winners'];

    for (const sheetName of sheetsToCheck) {
        try {
            const data = await getSheetData(sheetName);
            const records = parseSheetData(data);

            for (let i = 0; i < records.length; i++) {
                if (records[i]['User ID'] === userId.toString()) {
                    return { sheet: sheetName, rowIndex: i + 2, record: records[i] };
                }
            }
        } catch (error) {
            console.error(`Error searching in ${sheetName}:`, error);
        }
    }

    return { sheet: null, rowIndex: null, record: null };
}

client.once('ready', async () => {
    console.log(`Logged in as ${client.user.tag}!`);
    await initializeGoogleSheets();
    await registerCommands();
    await startCycleTimer();
});

client.on('guildCreate', async (guild) => {
    console.log(`Joined new guild: ${guild.name}`);
    try {
        const commands = await buildCommandsArray();
        await guild.commands.set(commands);
        console.log(`Registered commands for new guild: ${guild.name}`);
    } catch (error) {
        console.error(`Error registering commands for new guild ${guild.name}:`, error);
    }
});

function buildCommandsArray() {
    return [
        new SlashCommandBuilder()
            .setName('signup')
            .setDescription('Sign up for an election')
            .addStringOption(option =>
                option.setName('state')
                    .setDescription('The state you want to run in')
                    .setRequired(true)
                    .addChoices(
                        { name: 'Columbia', value: 'Columbia' },
                        { name: 'Cambridge', value: 'Cambridge' },
                        { name: 'Austin', value: 'Austin' },
                        { name: 'Superior', value: 'Superior' },
                        { name: 'Heartland', value: 'Heartland' },
                        { name: 'Yellowstone', value: 'Yellowstone' },
                        { name: 'Phoenix', value: 'Phoenix' }
                    ))
            .addStringOption(option =>
                option.setName('name')
                    .setDescription('Your candidate\'s name')
                    .setRequired(true))
            .addStringOption(option =>
                option.setName('party')
                    .setDescription('Your political party')
                    .setRequired(true)
                    .addChoices(
                        { name: 'Democrats', value: 'Democrats' },
                        { name: 'Republicans', value: 'Republicans' },
                        { name: 'Independent', value: 'Independent' }
                    )),
        new SlashCommandBuilder()
            .setName('withdraw')
            .setDescription('Withdraw from an election')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('The name of the character to withdraw')
                    .setRequired(true)),
        new SlashCommandBuilder()
            .setName('speech')
            .setDescription('Hold a speech to gain points (1 character = 1 point)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character giving the speech')
                    .setRequired(true)),
        new SlashCommandBuilder()
            .setName('canvassing')
            .setDescription('Go door-to-door canvassing for votes (0.5-1 point)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character (leave blank to use your own)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('donor')
            .setDescription('Accept donor funds (5 corruption, 3-6 points)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character (leave blank to use your own)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('special')
            .setDescription('Speech to special interest group (corruption and points per paragraph)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character giving the speech')
                    .setRequired(true))
            .addStringOption(option =>
                option.setName('speech')
                    .setDescription('The speech (1-5 paragraphs, 3 corruption and 2-4 points each)')
                    .setRequired(true)),
        new SlashCommandBuilder()
            .setName('ad')
            .setDescription('Submit a campaign video ad (1-3 points)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character (leave blank to use your own)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('poster')
            .setDescription('Submit a campaign poster image (0.5-1 point)')
            .addStringOption(option =>
                option.setName('character_name')
                    .setDescription('Name of the character (leave blank to use your own)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('regionpoll')
            .setDescription('Simulate a regional poll for Senate or Governor races')
            .addStringOption(option =>
                option.setName('office')
                    .setDescription('The office to poll')
                    .setRequired(true)
                    .addChoices(
                        { name: 'Senate', value: 'Senate' },
                        { name: 'Governor', value: 'Governor' }
                    ))
            .addStringOption(option =>
                option.setName('state')
                    .setDescription('The state to poll')
                    .setRequired(true)
                    .addChoices(
                        { name: 'Columbia', value: 'Columbia' },
                        { name: 'Cambridge', value: 'Cambridge' },
                        { name: 'Austin', value: 'Austin' },
                        { name: 'Superior', value: 'Superior' },
                        { name: 'Heartland', value: 'Heartland' },
                        { name: 'Yellowstone', value: 'Yellowstone' },
                        { name: 'Phoenix', value: 'Phoenix' }
                    ))
            .addStringOption(option =>
                option.setName('type')
                    .setDescription('Type of poll')
                    .setRequired(true)
                    .addChoices(
                        { name: 'General', value: 'general' },
                        { name: 'Primary', value: 'primary' }
                    ))
            .addStringOption(option =>
                option.setName('party')
                    .setDescription('The party to poll (only for primaries)')
                    .setRequired(false)
                    .addChoices(
                        { name: 'Republicans', value: 'Republicans' },
                        { name: 'Democrats', value: 'Democrats' },
                        { name: 'Independent', value: 'Independent' }
                    )),
        new SlashCommandBuilder()
            .setName('list_signups')
            .setDescription('List all candidates by phase, race, and party')
            .addStringOption(option =>
                option.setName('state')
                    .setDescription('The state to filter by')
                    .setRequired(true)
                    .addChoices(
                        { name: 'Columbia', value: 'Columbia' },
                        { name: 'Cambridge', value: 'Cambridge' },
                        { name: 'Austin', value: 'Austin' },
                        { name: 'Superior', value: 'Superior' },
                        { name: 'Heartland', value: 'Heartland' },
                        { name: 'Yellowstone', value: 'Yellowstone' },
                        { name: 'Phoenix', value: 'Phoenix' }
                    ))
            .addStringOption(option =>
                option.setName('seat')
                    .setDescription('The seat ID to filter by (optional)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('commands')
            .setDescription('Show all available commands organized by category'),
        new SlashCommandBuilder()
            .setName('tally_winners')
            .setDescription('[ADMIN] Tally up points and determine winners by seat and party')
            .setDefaultMemberPermissions(0),
        new SlashCommandBuilder()
            .setName('transfer_winners')
            .setDescription('[ADMIN] Transfer declared winners from All Signups to All Winners sheet')
            .setDefaultMemberPermissions(0),
        new SlashCommandBuilder()
            .setName('pause')
            .setDescription('[ADMIN] Pause or resume the cycle timer')
            .setDefaultMemberPermissions(0)
            .addBooleanOption(option =>
                option.setName('pause')
                    .setDescription('True to pause, False to resume')
                    .setRequired(true)),
        new SlashCommandBuilder()
            .setName('change_date')
            .setDescription('[ADMIN] Change the current year, cycle, or month')
            .setDefaultMemberPermissions(0)
            .addIntegerOption(option =>
                option.setName('year')
                    .setDescription('New year (optional)')
                    .setRequired(false))
            .addIntegerOption(option =>
                option.setName('cycle')
                    .setDescription('New cycle number (optional)')
                    .setRequired(false))
            .addIntegerOption(option =>
                option.setName('month')
                    .setDescription('New month (optional)')
                    .setRequired(false)),
        new SlashCommandBuilder()
            .setName('time')
            .setDescription('Display the current game time (year, month, cycle, and phase)')
    ];
}

async function registerCommands() {
    const commands = buildCommandsArray();

    try {
        console.log('Started refreshing application (/) commands');

        // Clear global commands first
        await client.application?.commands.set([]);

        // Register to all guilds only (for immediate availability)
        const guilds = client.guilds.cache;
        for (const guild of guilds.values()) {
            try {
                await guild.commands.set(commands);
                console.log(`Registered commands for guild: ${guild.name}`);
            } catch (guildError) {
                console.error(`Error registering commands for guild ${guild.name}:`, guildError);
            }
        }

        console.log('Successfully reloaded application (/) commands');
    } catch (error) {
        console.error('Error registering commands:', error);
    }
}

client.on('interactionCreate', async interaction => {
    if (!interaction.isChatInputCommand() && !interaction.isButton()) return;

    try {
        if (interaction.isChatInputCommand()) {
            switch (interaction.commandName) {
                case 'signup':
                    await handleSignup(interaction);
                    break;
                case 'withdraw':
                    await handleWithdraw(interaction);
                    break;
                case 'speech':
                    await handleSpeech(interaction);
                    break;
                case 'canvassing':
                    await handleCanvassing(interaction);
                    break;
                case 'donor':
                    await handleDonor(interaction);
                    break;
                case 'special':
                    await handleSpecial(interaction);
                    break;
                case 'ad':
                    await handleAd(interaction);
                    break;
                case 'poster':
                    await handlePoster(interaction);
                    break;
                case 'regionpoll':
                    await handleRegionPoll(interaction);
                    break;
                case 'list_signups':
                    await handleListSignups(interaction);
                    break;
                case 'commands':
                    await handleCommands(interaction);
                    break;
                case 'tally_winners':
                    await handleTallyWinners(interaction);
                    break;
                case 'transfer_winners':
                    await handleTransferWinners(interaction);
                    break;
                case 'pause':
                    await handlePause(interaction);
                    break;
                case 'change_date':
                    await handleChangeDate(interaction);
                    break;
                case 'time':
                    await handleTime(interaction);
                    break;
            }
        } else if (interaction.isButton()) {
            if (interaction.customId.startsWith('vote_')) {
                await handlePollVote(interaction);
            } else if (interaction.customId.startsWith('end_poll_')) {
                await handleEndPoll(interaction);
            } else if (interaction.customId.startsWith('seat_')) {
                await handleSeatSelection(interaction);
            }
        }
    } catch (error) {
        console.error('Error handling interaction:', error);
        if (!interaction.replied && !interaction.deferred) {
            await interaction.reply({ content: 'An error occurred while processing your request', ephemeral: true });
        }
    }
});

async function handleSignup(interaction) {
    await interaction.deferReply();

    try {
        const state = interaction.options.getString('state');
        const name = interaction.options.getString('name')?.trim();
        const party = interaction.options.getString('party');
        const userId = interaction.user.id;

        if (!name || name.length > 100) {
            await interaction.followUp({ content: 'Invalid candidate name. Must be 1-100 characters', ephemeral: true });
            return;
        }

        const { sheet, rowIndex, record } = await findUserInSheets(userId);
        if (record && (!record.Winner || !['withdrawn', 'loser'].includes(record.Winner.toLowerCase()))) {
            await interaction.followUp(`You are already signed up as '${record.Name}' for seat '${record['Seat ID']}'. Use \`/withdraw\` to cancel your candidacy first`);
            return;
        }

        const cycle = await getCurrentCycle();
        if (!cycle) {
            await interaction.editReply('Error accessing cycle data');
            return;
        }

        if (cycle.phase !== 'Signups') {
            await interaction.editReply('Signups are only allowed during the Signups phase');
            return;
        }

        const seats = await getEligibleSeats(cycle.year, state);
        if (seats.length === 0) {
            await interaction.editReply(`No eligible seats found for ${state} in ${cycle.year}`);
            return;
        }

        const maxSeatsToShow = Math.min(seats.length, 25);
        const components = [];

        for (let i = 0; i < maxSeatsToShow; i += 5) {
            const rowButtons = seats.slice(i, i + 5).map((seat, index) => 
                new ButtonBuilder()
                    .setCustomId(`seat_${i + index}_${interaction.id}`)
                    .setLabel(`${seat['Seat ID']} (${seat.Office})`)
                    .setStyle(ButtonStyle.Primary)
            );
            components.push(new ActionRowBuilder().addComponents(rowButtons));
        }

        const seatOptions = seats.map(s => `${s['Seat ID']} (${s.Office})`).join('\n');
        let content = `Available seats in ${state} for ${cycle.year}:\n${seatOptions}\n\nPlease click a button to select a seat:`;

        if (seats.length > 25) {
            content += `\n\n⚠️ Showing first 25 seats. Total seats available: ${seats.length}`;
        }

        const response = await interaction.editReply({ content, components });

        const signupData = new Map();
        signupData.set(interaction.id, { seats, state, name, party, userId, cycle });

        const filter = i => i.user.id === interaction.user.id && i.customId.startsWith(`seat_`) && i.customId.includes(interaction.id);
        const collector = response.createMessageComponentCollector({ filter, time: 60000 });

        collector.on('collect', async i => {
            await handleSeatSelection(i, signupData);
        });

        collector.on('end', async collected => {
            if (collected.size === 0) {
                const disabledComponents = components.map(row => {
                    const disabledRow = new ActionRowBuilder();
                    row.components.forEach(button => {
                        disabledRow.addComponents(ButtonBuilder.from(button).setDisabled(true));
                    });
                    return disabledRow;
                });
                await response.edit({ content: 'Seat selection timed out', components: disabledComponents });
            }
            signupData.delete(interaction.id);
        });
    } catch (error) {
        console.error('Error in signup:', error);
        await interaction.editReply(`Error processing signup: ${error.message}`);
    }
}

async function handleSeatSelection(interaction, signupData) {
    await interaction.deferUpdate();

    try {
        const [_, seatIndex, interactionId] = interaction.customId.split('_');
        const data = signupData.get(interactionId);
        if (!data) {
            await interaction.followUp({ content: 'Signup session expired', ephemeral: true });
            return;
        }

        const { seats, state, name, party, userId, cycle } = data;
        const selectedSeat = seats[parseInt(seatIndex)];

        if (!selectedSeat || seatIndex >= seats.length) {
            await interaction.followUp({ content: 'Invalid seat selection', ephemeral: true });
            return;
        }

        const cycleData = await getSheetData('Cycles');
        let stateFromCycles = state;
        for (const row of cycleData) {
            if (row[0] === selectedSeat['Seat ID']) {
                stateFromCycles = row[2] || state;
                break;
            }
        }

        const signupRow = [
            userId,
            name,
            selectedSeat['Seat ID'],
            party,
            cycle.phase,
            stateFromCycles,
            selectedSeat.Office,
            '0',
            '100',
            '0',
            ''
        ];

        await appendToSheet('All Signups', signupRow);

        await interaction.followUp({
            content: `${name} signed up for ${selectedSeat['Seat ID']} (${selectedSeat.Office}) in ${stateFromCycles} as ${party}!`,
            ephemeral: true
        });

        const disabledComponents = interaction.message.components.map(row => {
            const disabledRow = new ActionRowBuilder();
            row.components.forEach(button => {
                disabledRow.addComponents(ButtonBuilder.from(button).setDisabled(true));
            });
            return disabledRow;
        });

        await interaction.message.edit({ content: 'Seat selection complete!', components: disabledComponents });

        signupData.delete(interactionId);
    } catch (error) {
        console.error('Error in seat selection:', error);
        await interaction.followUp({ content: `Error processing seat selection: ${error.message}`, ephemeral: true });
    }
}

async function handleWithdraw(interaction) {
    await interaction.deferReply();

    try {
        const charName = interaction.options.getString('character_name')?.trim();
        const userId = interaction.user.id;

        if (!charName || charName.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters', ephemeral: true });
            return;
        }

        const data = await getSheetData('All Signups');
        const records = parseSheetData(data);

        for (let i = 0; i < records.length; i++) {
            const record = records[i];
            if (record.Name?.toLowerCase() === charName.toLowerCase() && 
                record['User ID'] === userId.toString()) {
                if (['withdrawn', 'winner'].includes(record.Winner?.toLowerCase())) {
                    await interaction.editReply(`Cannot withdraw '${charName}'. Candidacy is already marked as '${record.Winner}'`);
                    return;
                }

                await deleteSheetRow('All Signups', i + 2);
                await interaction.editReply(`${charName} has withdrawn their candidacy`);
                return;
            }
        }

        await interaction.editReply(`No signup found for '${charName}' associated with your account`);
    } catch (error) {
        console.error('Error during withdrawal:', error);
        await interaction.editReply(`Error during withdrawal: ${error.message}`);
    }
}

async function handleSpeech(interaction) {
    await interaction.deferReply({ ephemeral: true });

    try {
        const charName = interaction.options.getString('character_name')?.trim();

        if (!charName || charName.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters', ephemeral: true });
            return;
        }

        const promptMessage = await interaction.followUp({ 
            content: `Please reply with your speech text for ${charName} within 5 minutes. You'll earn 1 point for every character in your speech.`, 
            ephemeral: false 
        });

        const filterMessages = m => m.author.id === interaction.user.id && m.channel.id === interaction.channel.id;
        const collector = interaction.channel.createMessageCollector({ filter: filterMessages, time: 300000, max: 1 });

        collector.on('collect', async message => {
            const speech = message.content?.trim();

            if (!speech || speech.length > 3000) {
                await message.reply({ content: 'Invalid speech. Must be 1-3000 characters', ephemeral: true });
                return;
            }

            // Calculate points based on character count (1 character = 1 point, capped at 500)
            const points = Math.min(speech.length, 3000); // Cap at 500 points to prevent abuse

            if (points === 0) {
                await message.reply('Your speech needs at least 1 character to earn points');
                return;
            }

            await message.reply(`Processing speech for ${charName}... (${speech.length} characters = ${points} point${points !== 1 ? 's' : ''})`);

            await handleSpeechAction(interaction, charName, points, speech, message);
        });

        collector.on('end', async (collected) => {
            if (collected.size === 0) {
                await interaction.followUp({ content: 'No speech was submitted within 5 minutes. Speech submission cancelled.', ephemeral: true });
            }
        });
    } catch (error) {
        console.error('Error in speech:', error);
        await interaction.followUp({ content: `Error processing speech: ${error.message}`, ephemeral: true });
    }
}

async function handleSpeechAction(interaction, charName, points, speech, message) {
    try {
        const data = await getSheetData('All Signups');
        const records = parseSheetData(data);

        for (let i = 0; i < records.length; i++) {
            const record = records[i];
            if (record.Name?.toLowerCase() === charName.toLowerCase() && 
                record.Winner?.toLowerCase() !== 'withdrawn') {
                const stamina = parseInt(record.Stamina) || 0;
                const staminaCost = 10;

                if (stamina < staminaCost) {
                    await message.reply(`${record.Name} has insufficient stamina. Need ${staminaCost}, have ${stamina}`);
                    return;
                }

                const newStamina = Math.max(stamina - staminaCost, 0);
                const currentPoints = parseFloat(record.Points) || 0;
                const newPoints = currentPoints + points;

                await updateSheetCell('All Signups', i + 2, 8, newStamina);
                await updateSheetCell('All Signups', i + 2, 9, newPoints);

                let replyMsg = `${record.Name} held a speech and gained ${points} point${points !== 1 ? 's' : ''}! ` +
                    `Stamina: ${newStamina}, Points: ${newPoints}`;

                if (speech) {
                    replyMsg += `\n\n**Speech**: "${speech.substring(0, 300)}${speech.length > 300 ? '...' : ''}"`;
                }

                await message.reply(replyMsg);
                return;
            }
        }

        await message.reply(`No character named "${charName}" found in All Signups`);
    } catch (error) {
        console.error('Error in speech action:', error);
        await message.reply(`Error during speech processing: ${error.message}`);
    }
}

async function handleCanvassing(interaction) {
    await interaction.deferReply();

    try {
        const charName = interaction.options.getString('character_name')?.trim();

        if (charName?.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters or empty', ephemeral: true });
            return;
        }

        await handleCampaignAction(interaction, 'canvassing', 0, [0.5, 1], charName);
    } catch (error) {
        console.error('Error in canvassing:', error);
        await interaction.editReply(`Error processing canvassing: ${error.message}`);
    }
}

async function handleDonor(interaction) {
    await interaction.deferReply();

    try {
        const charName = interaction.options.getString('character_name')?.trim();

        if (charName?.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters or empty', ephemeral: true });
            return;
        }

        const corruptionGain = 5;
        const points = Math.floor(Math.random() * (6 - 3 + 1)) + 3;

        await handleCampaignAction(interaction, 'donor', 0, [points, points], charName, corruptionGain);
    } catch (error) {
        console.error('Error in donor:', error);
        await interaction.editReply(`Error processing donor: ${error.message}`);
    }
}

async function handleSpecial(interaction) {
    await interaction.deferReply();

    try {
        const charName = interaction.options.getString('character_name')?.trim();
        const speech = interaction.options.getString('speech')?.trim();

        if (!charName || charName.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters', ephemeral: true });
            return;
        }

        if (!speech || speech.length > 2000) {
            await interaction.followUp({ content: 'Invalid speech. Must be 1-2000 characters', ephemeral: true });
            return;
        }

        const paragraphs = speech.split(/\n\s*\n/).filter(p => p.trim().length >= 10);
        const numParagraphs = Math.min(paragraphs.length, 5);

        if (numParagraphs === 0) {
            await interaction.editReply('Your speech needs at least one paragraph (minimum 10 characters) to earn points');
            return;
        }

        const corruptionGain = numParagraphs * 3;
        let totalPoints = 0;
        for (let i = 0; i < numParagraphs; i++) {
            totalPoints += Math.floor(Math.random() * (4 - 2 + 1)) + 2;
        }

        await handleCampaignAction(interaction, 'special interest speech', 0, [totalPoints, totalPoints], charName, corruptionGain, speech);
    } catch (error) {
        console.error('Error in special:', error);
        await interaction.editReply(`Error processing special interest speech: ${error.message}`);
    }
}

async function handleAd(interaction) {
    await interaction.deferReply({ ephemeral: true });

    try {
        const charName = interaction.options.getString('character_name')?.trim();

        if (charName?.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters or empty', ephemeral: true });
            return;
        }

        const promptMessage = await interaction.followUp({ content: 'Please attach a video file (.mp4, .mov, .webm) in this channel within 60 seconds.', ephemeral: false });

        const filterMessages = m => m.author.id === interaction.user.id && m.attachments.size > 0 && m.channel.id === interaction.channel.id;
        const collector = interaction.channel.createMessageCollector({ filter: filterMessages, time: 60000, max: 1 });

        collector.on('collect', async message => {
            const attachment = message.attachments.first();
            const validExtensions = ['.mp4', '.mov', '.webm'];
            const isValidVideo = validExtensions.some(ext => attachment.name.toLowerCase().endsWith(ext));

            if (!isValidVideo) {
                await message.reply({ content: 'Invalid file type. Please attach a video file (.mp4, .mov, .webm)', ephemeral: true });
                return;
            }

            const points = Math.floor(Math.random() * (3 - 1 + 1)) + 1;

            await handleCampaignAction(interaction, 'campaign video ad', 0, [points, points], charName, 0, false, attachment.url);
            await message.reply({ content: `Campaign video ad submitted successfully! Earned ${points} point${points !== 1 ? 's' : ''}.` });
        });

        collector.on('end', async (collected) => {
            if (collected.size === 0) {
                await interaction.followUp({ content: 'No valid video file was attached within 60 seconds. Ad submission cancelled.', ephemeral: true });
            }
        });
    } catch (error) {
        console.error('Error in handleAd:', error);
        await interaction.followUp({ content: `Error processing campaign ad: ${error.message}`, ephemeral: true });
    }
}

async function handlePoster(interaction) {
    await interaction.deferReply({ ephemeral: true });

    try {
        const charName = interaction.options.getString('character_name')?.trim();

        if (charName?.length > 100) {
            await interaction.followUp({ content: 'Invalid character name. Must be 1-100 characters or empty', ephemeral: true });
            return;
        }

        const promptMessage = await interaction.followUp({ content: 'Please attach an image file (.png, .jpg, .jpeg, .gif) in this channel within 60 seconds.', ephemeral: false });

        const filterMessages = m => m.author.id === interaction.user.id && m.attachments.size > 0 && m.channel.id === interaction.channel.id;
        const collector = interaction.channel.createMessageCollector({ filter: filterMessages, time: 60000, max: 1 });

        collector.on('collect', async message => {
            const attachment = message.attachments.first();
            const validExtensions = ['.png', '.jpg', '.jpeg', '.gif'];
            const isValidImage = validExtensions.some(ext => attachment.name.toLowerCase().endsWith(ext));

            if (!isValidImage) {
                await message.reply({ content: 'Invalid file type. Please attach an image file (.png, .jpg, .jpeg, .gif)', ephemeral: true });
                return;
            }

            const points = Math.random() * (1 - 0.5) + 0.5;

            await handleCampaignAction(interaction, 'campaign poster', 0, [points, points], charName, 0, true, attachment.url);
            await message.reply({ content: `Campaign poster submitted successfully! Earned ${points.toFixed(2)} point${points !== 1 ? 's' : ''}.` });
        });

        collector.on('end', async (collected) => {
            if (collected.size === 0) {
                await interaction.followUp({ content: 'No valid image file was attached within 60 seconds. Poster submission cancelled.', ephemeral: true });
            }
        });
    } catch (error) {
        console.error('Error in handlePoster:', error);
        await interaction.followUp({ content: `Error processing campaign poster: ${error.message}`, ephemeral: true });
    }
}

async function handleCampaignAction(interaction, actionName, staminaCost, pointsRange, charName, corruptionGain = 0, extraInfo = '', isImage = false) {
    try {
        const userId = interaction.user.id;
        const cycle = await getCurrentCycle();
        if (!cycle) {
            await interaction.reply({ content: 'Error accessing cycle data', ephemeral: true });
            return;
        }

        const sheetName = cycle.phase === 'General Election' || cycle.phase === 'General Campaign' ? 'All Winners' : 'All Signups';
        console.log(`Handling ${actionName} for phase ${cycle.phase}, using sheet ${sheetName}`);

        await createSheetIfNotExists(sheetName);
        const data = await getSheetData(sheetName);
        const records = parseSheetData(data);

        for (let i = 0; i < records.length; i++) {
            const record = records[i];
            const isMatch = charName
                ? record.Name?.toLowerCase() === charName.toLowerCase() && 
                  record['User ID'] === userId.toString()
                : record['User ID'] === userId.toString();

            if (isMatch && record.Winner?.toLowerCase() !== 'withdrawn') {
                const stamina = parseInt(record.Stamina) || 0;
                const currentCorruption = parseInt(record.Corruption) || 0;
                const currentPoints = parseFloat(record.Points) || 0;

                if (staminaCost > 0 && stamina < staminaCost) {
                    await interaction.reply({ content: `${record.Name} has insufficient stamina. Need ${staminaCost}, have ${stamina}`, ephemeral: true });
                    return;
                }

                const points = actionName === 'canvassing' || actionName === 'campaign poster'
                    ? Math.random() * (pointsRange[1] - pointsRange[0]) + pointsRange[0]
                    : Math.floor(Math.random() * (pointsRange[1] - pointsRange[0] + 1)) + pointsRange[0];

                const newStamina = staminaCost > 0 ? Math.max(stamina - staminaCost, 0) : stamina;
                const newCorruption = Math.min(currentCorruption + corruptionGain, 100);
                const newPoints = currentPoints + points;

                if (staminaCost > 0) {
                    await updateSheetCell(sheetName, i + 2, 8, newStamina.toString());
                }
                if (corruptionGain > 0) {
                    await updateSheetCell(sheetName, i + 2, 7, newCorruption.toString());
                }
                await updateSheetCell(sheetName, i + 2, 9, newPoints.toFixed(2));

                let replyMsg = `${record.Name} performed ${actionName} and gained ${points.toFixed(2)} point${points !== 1 ? 's' : ''}! `;
                if (staminaCost > 0) replyMsg += `Stamina: ${newStamina}, `;
                if (corruptionGain > 0) replyMsg += `Corruption: ${newCorruption}, `;
                replyMsg += `Points: ${newPoints.toFixed(2)}`;
                if (extraInfo) {
                    replyMsg += `\n\n**${actionName.charAt(0).toUpperCase() + actionName.slice(1)}**: `;
                    if (actionName.includes('speech')) {
                        replyMsg += `"${extraInfo.substring(0, 200)}${extraInfo.length > 200 ? '...' : ''}"`;
                    } else if (isImage) {
                        replyMsg += `![Image](${extraInfo})`;
                    } else {
                        replyMsg += `[Video](${extraInfo})`;
                    }
                }

                await interaction.followUp({ content: replyMsg, ephemeral: true });
                return;
            }
        }

        await interaction.followUp({ content: `No active candidacy found for ${charName ? `character "${charName}"` : 'you'} in ${sheetName}. Please sign up first.`, ephemeral: true });
    } catch (error) {
        console.error(`Error in ${actionName}:`, error);
        await interaction.followUp({ content: `Error during ${actionName}: ${error.message}`, ephemeral: true });
    }
}

function generateTimeProgressBar(startTime) {
    const totalDurationMs = 24 * 60 * 60 * 1000; // 24 hours
    const elapsedMs = Date.now() - startTime;
    const progress = Math.min(elapsedMs / totalDurationMs, 1);
    const segments = 10;
    const filledSegments = Math.round(progress * segments);
    const emptySegments = segments - filledSegments;
    const remainingHours = Math.floor((totalDurationMs - elapsedMs) / (60 * 60 * 1000));
    const remainingMinutes = Math.floor((totalDurationMs - elapsedMs) % (60 * 60 * 1000) / (60 * 1000));
    return `${'█'.repeat(filledSegments)}${'-'.repeat(emptySegments)} (${remainingHours}h ${remainingMinutes}m left)`;
}

async function updatePollEmbed(pollMessage, candidates, pollId, cycle, office, state, party, sheetName) {
    try {
        const { startTime } = pollData.get(pollId);
        if (!startTime) {
            console.warn(`No start time found for poll ${pollId}`);
            return;
        }

        const isPrimary = cycle.phase.toLowerCase().includes('primary') || cycle.phase.toLowerCase() === 'signups';
        const pollType = isPrimary ? 'Primary' : 'General';
        const partyText = party ? ` - ${party}` : (isPrimary ? ' - All Parties' : '');

        const embed = new EmbedBuilder()
            .setTitle(`📊 ${pollType} Poll for ${office} in ${state}${partyText}`)
            .setColor(0x0099ff)
            .setDescription(
                candidates.map((c, i) => `${i + 1}. ${c.Name} (${c['Seat ID']}) - ${c.Party}`).join('\n') +
                '\n\nClick a candidate to vote (once only).\n' +
                `**Time Remaining**: ${generateTimeProgressBar(startTime)}`
            )
            .setFooter({ text: `${pollType} Phase: ${cycle.phase} | Source: ${sheetName} | Poll ends at ${new Date(startTime + 24 * 60 * 60 * 1000).toLocaleString()}` });

        await pollMessage.edit({ embeds: [embed], components: pollMessage.components });
    } catch (error) {
        console.error(`Error updating poll embed for ${pollId}:`, error);
    }
}

async function endPoll(pollMessage, pollId, candidates, cycle, office, state, party, reason, sheetName) {
    try {
        const { intervalId, collector } = pollData.get(pollId);
        if (intervalId) {
            clearInterval(intervalId);
            console.log(`Cleared interval for poll ${pollId}`);
        }
        if (collector) collector.stop();

        const disabledComponents = pollMessage.components.map(row => {
            const disabledRow = new ActionRowBuilder();
            row.components.forEach(button => {
                disabledRow.addComponents(ButtonBuilder.from(button).setDisabled(true));
            });
            return disabledRow;
        });

        await pollMessage.edit({ components: disabledComponents });

        const votes = pollVotes.get(pollId) || new Map();
        const voteCounts = Array(candidates.length).fill(0);
        for (const [, voteIndex] of votes) {
            voteCounts[voteIndex]++;
        }

        const totalVotes = voteCounts.reduce((a, b) => a + b, 0);

        const votePercentages = voteCounts.map(count => 
            totalVotes > 0 ? (count / totalVotes) * 100 : 0
        );

        const voteProgressBars = votePercentages.map(percentage => {
            const filled = Math.round(percentage / 10);
            const empty = 10 - filled;
            return `${'█'.repeat(filled)}${'-'.repeat(empty)} ${percentage.toFixed(1)}%`;
        });

        const maxVotes = Math.max(...voteCounts);
        const winners = candidates.filter((c, i) => voteCounts[i] === maxVotes && maxVotes > 0);
        const winnerNames = winners.map(c => c.Name);

        let resultMsg = '';
        if (winners.length > 0) {
            const data = await getSheetData(sheetName);
            const records = parseSheetData(data);
            const winnersUpdated = [];

            for (let i = 0; i < records.length; i++) {
                const record = records[i];
                if (winnerNames.includes(record.Name) && 
                    record.Office?.toLowerCase() === office.toLowerCase() && 
                    record.States?.toLowerCase() === state.toLowerCase() && 
                    (!party || record.Party?.toLowerCase() === party?.toLowerCase()) && 
                    record.Winner?.toLowerCase() !== 'withdrawn') {
                    const currentPoints = parseFloat(record.Points) || 0;
                    const newPoints = currentPoints + 8;
                    await updateSheetCell(sheetName, i + 2, 9, newPoints.toFixed(2));
                    winnersUpdated.push(`${record.Name} (${record['Seat ID']})`);
                }
            }

            resultMsg = `**Winner(s):** ${winnersUpdated.join(', ')}\n\nEach winner received 8 points\n\n` +
                candidates.map((c, i) => 
                    `${c.Name} (${c['Seat ID']}): ${voteCounts[i]} vote${voteCounts[i] !== 1 ? 's' : ''}\n${voteProgressBars[i]}`
                ).join('\n\n');
        } else {
            resultMsg = `Poll for ${office} in ${state}${party ? ` (${party})` : ''} ended with no votes`;
        }

        const resultEmbed = new EmbedBuilder()
            .setTitle(`📊 Poll Results for ${office} in ${state}${party ? ` (${party})` : ''}`)
            .setColor(0x00ff00)
            .setDescription(resultMsg)
            .setFooter({ text: `Phase: ${cycle.phase} | Source: ${sheetName} | Total Votes: ${totalVotes} | Ended: ${reason}` });

        await pollMessage.channel.send({ embeds: [resultEmbed] });

        pollVotes.delete(pollId);
        pollData.delete(pollId);
    } catch (error) {
        console.error(`Error ending poll ${pollId}:`, error);
        await pollMessage.channel.send(`Error ending poll: ${error.message}`);
        pollVotes.delete(pollId);
        pollData.delete(pollId);
    }
}

async function handleRegionPoll(interaction) {
    await interaction.deferReply();

    try {
        const office = interaction.options.getString('office');
        const state = interaction.options.getString('state');
        const pollType = interaction.options.getString('type');
        const party = interaction.options.getString('party');
        
        // Validate party requirement for primaries
        if (pollType === 'primary' && !party) {
            await interaction.editReply('Party selection is required for primary polls');
            return;
        }

        const cycle = await getCurrentCycle();
        if (!cycle) {
            await interaction.editReply('Error accessing cycle data');
            return;
        }

        // Determine data source based on poll type
        const sheetName = pollType === 'general' ? 'All Winners' : 'All Signups';
        console.log(`Creating simulated poll: Office=${office}, State=${state}, Type=${pollType}, Party=${party || 'All'}, Source=${sheetName}`);

        const data = await getSheetData(sheetName);
        if (!data || data.length === 0) {
            await interaction.editReply(`No data available in ${sheetName}. ${pollType === 'general' ? 'Run primaries and transfer winners first' : 'No candidates registered'}`);
            return;
        }

        const records = parseSheetData(data);

        let candidates = records.filter(r => {
            const matchesOffice = r.Office?.toLowerCase() === office.toLowerCase();
            const matchesState = (r.States?.toLowerCase() === state.toLowerCase()) || (r.State?.toLowerCase() === state.toLowerCase());
            const matchesParty = !party || r.Party?.toLowerCase() === party?.toLowerCase();
            const notWithdrawn = r.Winner?.toLowerCase() !== 'withdrawn';

            // For All Signups sheet, also check if they're not marked as losers
            if (sheetName === 'All Signups') {
                const notLoser = r.Winner?.toLowerCase() !== 'loser';
                return matchesOffice && matchesState && matchesParty && notWithdrawn && notLoser;
            }

            return matchesOffice && matchesState && matchesParty && notWithdrawn;
        });

        if (candidates.length === 0) {
            let errorMsg = `No candidates found for ${office} in ${state}`;
            if (pollType === 'primary' && party) {
                errorMsg += ` (${party} primary)`;
            } else if (pollType === 'general') {
                errorMsg += ` (general election)`;
            }
            await interaction.editReply(errorMsg);
            return;
        }

        // Calculate poll scores for each candidate
        const candidateScores = candidates.map(candidate => {
            const points = parseFloat(candidate.Points) || 0;
            const stamina = parseFloat(candidate.Stamina) || 0;
            const corruption = parseFloat(candidate.Corruption) || 0;
            
            // Score = points + stamina/10 - corruption/10 (corruption is negative)
            const score = points + (stamina / 10) - (corruption / 10);
            
            return {
                name: candidate.Name || candidate.Candidate,
                party: candidate.Party,
                seatId: candidate['Seat ID'],
                score: Math.max(score, 0.1) // Ensure minimum score for percentage calculation
            };
        });

        // Calculate total score for percentage calculation
        const totalScore = candidateScores.reduce((sum, candidate) => sum + candidate.score, 0);

        // Calculate percentages
        const candidateResults = candidateScores.map(candidate => ({
            ...candidate,
            percentage: (candidate.score / totalScore) * 100
        }));

        // Sort by percentage (highest first)
        candidateResults.sort((a, b) => b.percentage - a.percentage);

        // Generate random total votes between 10,000-100,000
        const totalVotes = Math.floor(Math.random() * (100000 - 10000 + 1)) + 10000;

        // Calculate actual vote counts based on percentages
        const candidateVotes = candidateResults.map((candidate, index) => {
            let votes;
            if (index === candidateResults.length - 1) {
                // Last candidate gets remaining votes to ensure total is exact
                const usedVotes = candidateResults.slice(0, -1).reduce((sum, c) => sum + (c.votes || 0), 0);
                votes = totalVotes - usedVotes;
            } else {
                votes = Math.round((candidate.percentage / 100) * totalVotes);
            }
            
            return {
                ...candidate,
                votes: Math.max(votes, 0)
            };
        });

        const pollTypeName = pollType === 'primary' ? 'Primary' : 'General';
        const partyText = pollType === 'primary' && party ? ` - ${party}` : (pollType === 'general' ? ' - All Parties' : '');

        // Create results description
        const resultsDescription = candidateVotes.map((candidate, index) => {
            const progressBar = '█'.repeat(Math.round(candidate.percentage / 10)) + 
                               '-'.repeat(10 - Math.round(candidate.percentage / 10));
            return `${index + 1}. **${candidate.name}** (${candidate.party})\n` +
                   `   ${progressBar} ${candidate.percentage.toFixed(1)}% (${candidate.votes.toLocaleString()} votes)`;
        }).join('\n\n');

        const embed = new EmbedBuilder()
            .setTitle(`📊 Regional ${pollTypeName} Poll for ${office} in ${state}${partyText}`)
            .setColor(0x00ff00)
            .setDescription(
                `**Poll Results** (${totalVotes.toLocaleString()} total votes)\n\n` +
                resultsDescription
            )
            .setFooter({ text: `${pollTypeName} Poll | Source: ${sheetName} | Simulated at ${new Date().toLocaleString()}` });

        await interaction.editReply({ embeds: [embed] });

    } catch (error) {
        console.error('Error in region poll creation:', error);
        await interaction.editReply(`Error creating simulated poll: ${error.message}`);
    }
}

async function handlePollVote(interaction) {
}

async function handleEndPoll(interaction) {
    const pollId = interaction.message.id;
    if (!pollData.has(pollId)) {
        await interaction.reply({ content: 'This poll has already ended or is invalid', ephemeral: true });
        return;
    }

    const { creatorId } = pollData.get(pollId);
    if (interaction.user.id !== creatorId && !interaction.member.permissions.has('Administrator')) {
        await interaction.reply({ content: 'Only the poll creator or an admin can end this poll', ephemeral: true });
        return;
    }

    await interaction.deferReply({ ephemeral: true });

    const { collector } = pollData.get(pollId);
    collector.stop();
    await interaction.followUp({ content: 'Poll ended manually', ephemeral: true });
}

async function handleCommands(interaction) {
    await interaction.deferReply();

    try {
        const embed = new EmbedBuilder()
            .setTitle('🗳️ APRP Election Bot Commands')
            .setColor(0x0099ff)
            .addFields([
                {
                    name: '🎮 Player Commands',
                    value: [
                        '`/signup` - Sign up for an election',
                        '`/withdraw` - Withdraw from an election',
                        '`/speech` - Hold a speech to earn points (1 character = 1 point, costs 10 stamina)',
                        '`/canvassing` - Go door-to-door canvassing (0.5-1 point)',
                        '`/donor` - Accept donor funds (5 corruption, 3-6 points)',
                        '`/special` - Submit a special interest group speech (3 corruption and 2-4 points per paragraph)',
                        '`/ad` - Submit a campaign video ad (1-3 points)',
                        '`/poster` - Submit a campaign poster image (0.5-1 point)',
                        '`/regionpoll` - Simulate a regional poll for Senate or Governor races based on candidate stats',
                        '`/list_signups` - List all candidates by phase, race, and party',
                        '`/time` - Display the current game time (year, month, cycle, and phase)'
                    ].join('\n'),
                    inline: false
                },
                {
                    name: '🔧 Admin Commands',
                    value: [
                        '`/tally_winners` - Tally up points and determine winners by seat and party',
                        '`/transfer_winners` - Transfer declared winners to All Winners sheet',
                        '`/pause` - Pause or resume the cycle timer',
                        '`/change_date` - Change the current year, cycle, or month'
                    ].join('\n'),
                    inline: false
                },
                {
                    name: 'ℹ️ General Info',
                    value: '`/commands` - Show this help menu',
                    inline: false
                }
            ])
            .setFooter({ text: '💡 Tip: Manage corruption to avoid penalties!' });

        await interaction.editReply({ embeds: [embed] });
    } catch (error) {
        console.error('Error handling commands:', error);
        await interaction.editReply(`Error displaying commands: ${error.message}`);
    }
}

async function handleTallyWinners(interaction) {
    if (!interaction.member.permissions.has('Administrator')) {
        await interaction.reply({ content: 'You need administrator permissions to use this command!', ephemeral: true });
        return;
    }

    await interaction.deferReply();

    try {
        const data = await getSheetData('All Signups');
        const records = parseSheetData(data);
        console.log(`Tallying winners: Read ${records.length} records from All Signups`);

        const races = {};
        let skippedRecords = 0;

        records.forEach(record => {
            if (!record['User ID'] || !record.Name || !record['Seat ID'] || !record.Party) {
                console.log(`Skipping record: Missing required fields (User ID: ${record['User ID']}, Name: ${record.Name}, Seat ID: ${record['Seat ID']}, Party: ${record.Party})`);
                skippedRecords++;
                return;
            }

            if (['withdrawn', 'winner', 'loser'].includes(record.Winner?.toLowerCase())) {
                console.log(`Skipping record: Already processed (Name: ${record.Name}, Winner: ${record.Winner})`);
                skippedRecords++;
                return;
            }

            const key = `${record['Seat ID']}_${record.Party}`;
            if (!races[key]) {
                races[key] = [];
            }
            races[key].push(record);
            console.log(`Added to race ${key}: ${record.Name} (${record.Points} points)`);
        });

        let winnersCount = 0;

        for (const [key, candidates] of Object.entries(races)) {
            if (candidates.length === 0) {
                console.log(`No candidates for race ${key}, skipping`);
                continue;
            }

            console.log(`Processing race ${key} with ${candidates.length} candidates`);

            const winner = candidates.reduce((prev, current) => 
                (parseFloat(current.Points) || 0) > (parseFloat(prev.Points) || 0) ? current : prev
            );

            for (let i = 0; i < records.length; i++) {
                const record = records[i];
                if (record['Seat ID'] === winner['Seat ID'] && 
                    record.Party === winner.Party &&
                    !['withdrawn', 'winner', 'loser'].includes(record.Winner?.toLowerCase())) {

                    const status = record.Name === winner.Name ? 'Winner' : 'Loser';
                    await updateSheetCell('All Signups', i + 2, 10, status);
                    console.log(`Updated ${record.Name} in race ${key} to ${status}`);

                    if (status === 'Winner') {
                        winnersCount++;
                    }
                }
            }
        }

        console.log(`Tally complete: ${winnersCount} winners declared, ${skippedRecords} records skipped`);
        await interaction.editReply(`Tallying complete! ${winnersCount} winners declared across all seats and parties. Skipped ${skippedRecords} invalid or processed records.`);
    } catch (error) {
        console.error('Error in tally_winners:', error);
        await interaction.editReply(`Error tallying winners: ${error.message}`);
    }
}

async function handleTransferWinners(interaction) {
    if (!interaction.member.permissions.has('Administrator')) {
        await interaction.reply({ content: 'You need administrator permissions to use this command', ephemeral: true });
        return;
    }

    await interaction.deferReply();

    try {
        const signupsData = await getSheetData('All Signups');
        const winnersData = await getSheetData('All Winners');
        const cyclesData = await getSheetData('Cycles');

        let year = '1991';
        if (cyclesData && cyclesData[1] && cyclesData[1][6]) {
            year = cyclesData[1][6].toString();
        }

        if (winnersData.length === 0) {
            const headers = ["Year", "Office", "State", "Seat ID", "Candidate", "Party", "Points", "Votes", "Corruption", "Final Score", "Winner"];
            await appendToSheet('All Winners', headers);
        }

        const signupsRecords = parseSheetData(signupsData);
        const winnersRecords = parseSheetData(winnersData);

        let transferred = 0;

        for (const record of signupsRecords) {
            if (record.Winner?.toLowerCase() === 'winner') {
                const exists = winnersRecords.some(w => 
                    w['Seat ID'] === record['Seat ID'] && 
                    w.Candidate === record.Name &&
                    w.Year === year
                );

                if (!exists) {
                    const winnerRow = [
                        year,
                        record.Office || '',
                        record.States || '',
                        record['Seat ID'] || '',
                        record.Name || '',
                        record.Party || '',
                        '0',
                        '',
                        record.Corruption || '0',
                        '',
                        ''
                    ];

                    await appendToSheet('All Winners', winnerRow);
                    transferred++;
                }
            }
        }

        await interaction.editReply(`Transfer complete! ${transferred} winners transferred to All Winners sheet`);
    } catch (error) {
        console.error('Error in transfer_winners:', error);
        await interaction.editReply(`Error transferring winners: ${error.message}`);
    }
}

async function handlePause(interaction) {
    if (!interaction.member.permissions.has('Administrator')) {
        await interaction.reply({ content: 'You need administrator permissions to use this command!', ephemeral: true });
        return;
    }

    await interaction.deferReply();

    try {
        const pause = interaction.options.getBoolean('pause');
        await updateCyclePause(pause);
        await interaction.editReply(`Cycle timer is now ${pause ? 'paused' : 'resumed'}.`);
    } catch (error) {
        console.error('Error in pause:', error);
        await interaction.editReply(`Error pausing/resuming cycle: ${error.message}`);
    }
}

async function handleChangeDate(interaction) {
    if (!interaction.member.permissions.has('Administrator')) {
        await interaction.reply({ content: 'You need administrator permissions to use this command!', ephemeral: true });
        return;
    }

    await interaction.deferReply();

    try {
        const year = interaction.options.getInteger('year');
        const cycle = interaction.options.getInteger('cycle');
        const month = interaction.options.getInteger('month');

        if (year && (year < 1990 || year > 2100)) throw new Error('Year must be between 1990 and 2100');
        if (cycle && cycle < 1) throw new Error('Cycle must be positive');
        if (month && (month < 1 || month > 12)) throw new Error('Month must be between 1 and 12');

        if (year) await updateCycleYear(year);
        if (cycle) await updateCycleCycle(cycle);
        if (month) await updateCycleMonth(month);

        const newCycle = await getCurrentCycle();
        await interaction.editReply(`Date updated to Year: ${newCycle.year}, Cycle: ${newCycle.cycle}, Month: ${newCycle.month}`);
    } catch (error) {
        console.error('Error in change_date:', error);
        await interaction.editReply(`Error changing date: ${error.message}`);
    }
}

async function handleListSignups(interaction) {
    await interaction.deferReply();
    try {
        const state = interaction.options.getString('state');
        const seat = interaction.options.getString('seat');
        const data = await getSheetData('All Signups');
        const records = parseSheetData(data);

        if (!records || records.length === 0) {
            await interaction.editReply(`No candidates found in All Signups for ${state}${seat ? ` in seat ${seat}` : ''}.`);
            return;
        }

        let filteredRecords = records.filter(r => 
            r.States?.toLowerCase() === state.toLowerCase() && 
            (!seat || r['Seat ID'] === seat)
        );

        if (filteredRecords.length === 0) {
            await interaction.editReply(`No candidates found for ${state}${seat ? ` in seat ${seat}` : ''}.`);
            return;
        }

        const groupedByPhase = {};
        filteredRecords.forEach(record => {
            if (!groupedByPhase[record.Phase]) groupedByPhase[record.Phase] = {};
            if (!groupedByPhase[record.Phase][record['Seat ID']]) groupedByPhase[record.Phase][record['Seat ID']] = {};
            if (!groupedByPhase[record.Phase][record['Seat ID']][record.Party]) 
                groupedByPhase[record.Phase][record['Seat ID']][record.Party] = [];
            groupedByPhase[record.Phase][record['Seat ID']][record.Party].push(record.Name);
        });

        let content = `Candidates in ${state}${seat ? ` for ${seat}` : ''}:\n\n`;
        for (const phase in groupedByPhase) {
            content += `**${phase}**\n`;
            for (const seatId in groupedByPhase[phase]) {
                content += `  - ${seatId}\n`;
                for (const party in groupedByPhase[phase][seatId]) {
                    content += `    - ${party}: ${groupedByPhase[phase][seatId][party].join(', ')}\n`;
                }
            }
            content += '\n';
        }

        await interaction.editReply(content.length > 2000 ? content.substring(0, 1997) + '...' : content);
    } catch (error) {
        console.error('Error in list_signups:', error);
        await interaction.editReply(`Error listing signups: ${error.message}`);
    }
}

async function handleTime(interaction) {
    await interaction.deferReply();

    try {
        const cycle = await getCurrentCycle();
        if (!cycle) {
            await interaction.editReply('Error accessing cycle data');
            return;
        }

        const monthNames = [
            'January', 'February', 'March', 'April', 'May', 'June',
            'July', 'August', 'September', 'October', 'November', 'December'
        ];
        const monthName = monthNames[cycle.month - 1] || 'Unknown';

        const embed = new EmbedBuilder()
            .setTitle('🕰️ Current Game Time')
            .setColor(0x00ff00)
            .addFields([
                { name: 'Year', value: cycle.year.toString(), inline: true },
                { name: 'Month', value: monthName, inline: true },
                { name: 'Cycle', value: cycle.cycle.toString(), inline: true },
                { name: 'Phase', value: cycle.phase, inline: true },
                { name: 'Status', value: cycle.isPaused ? 'Paused' : 'Running', inline: true }
            ])
            .setTimestamp();

        await interaction.editReply({ embeds: [embed] });
    } catch (error) {
        console.error('Error in handleTime:', error);
        await interaction.editReply(`Error displaying time: ${error.message}`);
    }
}

// Start the cycle timer
async function startCycleTimer() {
    setInterval(async () => {
        try {
            const cycle = await getCurrentCycle();
            if (cycle.isPaused) return;

            let newMonth = cycle.month + 1;
            let newYear = cycle.year;
            let newCycle = cycle.cycle;

            if (newMonth > 12) {
                newMonth = 1;
                newYear++;
                newCycle++;
            }

            await updateCycleMonth(newMonth);
            await updateCycleYear(newYear);
            await updateCycleCycle(newCycle);

            // Update phase based on month and year
            const phases = [
                { name: 'Signups', start: 4, end: 8 },
                { name: 'Primary Campaign', start: 8, end: 12 },
                { name: 'Primary Election', start: 12, end: 12 },
                { name: 'General Campaign', start: 4, end: 9 },
                { name: 'General Election', start: 12, end: 12 }
            ];

            let newPhase = cycle.phase;
            for (const phase of phases) {
                if (newMonth >= phase.start && newMonth <= phase.end && (newYear === 1991 || newYear === 1992)) {
                    newPhase = phase.name;
                    break;
                }
            }

            if (newPhase !== cycle.phase) {
                await updateCyclePhase(newPhase);
            }
        } catch (error) {
            console.error('Error in cycle timer:', error);
        }
    }, 24 * 60 * 60 * 1000); // Run every 24 hours
}

// Login to Discord
client.login(process.env.DISCORD_TOKEN).catch(error => {
    console.error('Error logging in:', error);
    process.exit(1);
});


