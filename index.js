
// Loads environment variables from your .env file
require('dotenv').config();

// Imports the Discord bot client and the intents you need
// Client lets you create and control your bot.
// GatewayIntentBits tells Discord what kind of events your bot wants to listen to (like messages or guild joins).
const { Client, GatewayIntentBits } = require('discord.js');

const XLSX = require('xlsx');

//  auto-copies assest tag afte sending it
const clipboardy = require('clipboardy');

// Creates your bot instance.
// intents array tells Discord: 
//      Guilds: you want to connect to servers.
//      GuildMessages: you want to read messages in servers.
//      MessageContent: you want to read the actual text of messages (like /next_audit)
const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent] });

// Tracks which asset you’re currently on.
let assetIndex = 0;

// Gets filled when you call loadAssets()
let assetList = [];

function loadAssets() {

// Reads your Excel file.
  const workbook = XLSX.readFile('audit_data.xlsx');

  const sheet = workbook.Sheets[workbook.SheetNames[0]];

//Converts the first sheet into a JSON array.
  assetList = XLSX.utils.sheet_to_json(sheet);
}

// Returns the next asset from the list.
function getNextAsset() {
  if (assetIndex >= assetList.length) return null;
  const asset = assetList[assetIndex];
  assetIndex++;
  clipboardy.writeSync(asset.AssetTag); // Auto-copy to clipboard
  return `🗂️ Asset Tag: **${asset.AssetTag}**\n📅 Last Audit: ${asset.LastAuditDate}`;
}

// Loads your Excel data so the bot is ready to serve assets
client.once('ready', () => {

//Logs a message to your termina
  console.log(`Logged in as ${client.user.tag}`);
  loadAssets();
});

// If someone types /next_audit, it:
//      Gets the next asset.
//      Sends it as a message.
//      Copies the tag to your clipboard
client.on('messageCreate', async message => {
  if (message.content === '/next_audit') {
    const assetMsg = getNextAsset();
    if (!assetMsg) {
      message.channel.send('✅ All assets have been processed!');
    } else {
      message.channel.send(assetMsg + '\n📋 Copied to clipboard. Paste it into AssetSonar.');
    }
  }
});


// Logs your bot into Discord using the token from .env.
// Without this, your bot won’t connect.
client.login(process.env.DISCORD_TOKEN);
