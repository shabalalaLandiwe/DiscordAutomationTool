require('dotenv').config();
const { Client, GatewayIntentBits } = require('discord.js');
const XLSX = require('xlsx');
const clipboardy = require('clipboardy');

const client = new Client({ intents: [GatewayIntentBits.Guilds, GatewayIntentBits.GuildMessages, GatewayIntentBits.MessageContent] });

let assetIndex = 0;
let assetList = [];

function loadAssets() {
  const workbook = XLSX.readFile('audit_data.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  assetList = XLSX.utils.sheet_to_json(sheet);
}

function getNextAsset() {
  if (assetIndex >= assetList.length) return null;
  const asset = assetList[assetIndex];
  assetIndex++;
  clipboardy.writeSync(asset.AssetTag); // Auto-copy to clipboard
  return `🗂️ Asset Tag: **${asset.AssetTag}**\n📅 Last Audit: ${asset.LastAuditDate}`;
}

client.once('ready', () => {
  console.log(`Logged in as ${client.user.tag}`);
  loadAssets();
});

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

client.login(process.env.DISCORD_TOKEN);
