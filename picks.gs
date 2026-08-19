const VERSION = '1.2.1';
/** GOOGLE SHEETS FOOTBALL PICK 'EMS, SURVIVOR, & ELIMINATOR TOOL | 2025 Edition
 * Script Library for League Creator & Management Platform
 * 08/19/2026
 * 
 * Created by Ben Powers
 * ben.powers.creative@gmail.com
 * 
 * ------------------------------------------------------------------------
 * DESCRIPTION:
 * A series of Google Apps Scripts that generate multiple sheets and a weekly Google Form
 * to be utilized for gathering pick 'ems, survivor, and eliminator selections.
 * 
 * ------------------------------------------------------------------------
 * INSTRUCTIONS:
 * make a copy and try running the Picks menu option; if you have trouble authorizing the script, please follow the instructions; if you have no menu, please refresh
 * 
 * ------------------------------------------------------------
 * 🏈 PICKS MENU OPTIONS WITH FUNCTION EXPLANATIONS:
 * 
 *  📝 Form Builder - make a new form with all sorts of customization
 *  📋 Form Manager - review existing forms, turn on and off trigger to record logging, review specs of the form, copy links, etc. Also preview response count
 *  📥 Check & Import Responses - import picks for any week that has a form (only do this when you're ready to import)
 * 
 *  ⚙️ Configuration - set name, select which pools to run, and many other options
 * 
 *  👥 Member Manager - enter member names, rearrange, mark paid, revive (if using survivor/eliminator), and remove
 *  👑&💀 Survivor/Eliminator Manager - Allows for adjustments to the survivor and eliminator contests (only visible if Survivor or Eliminator Present)
 * 
 *  -----------
 * 
 *  🏈 Fetch Scores - bring in NFL outcomes to the sheet
 *  🧰 UTILITIES:
 *    📅 Update NFL Data - update the schedule data, should bring in new spreads
 *    📊 Update Spread Data - attempt to pull spreads for week, may need to wait to get updated data
 *    ✏️ Rename a Member - rename a member in the sheet and update back-end name (note: this won't update the name on the form, which could cause problems, do this mid-week)  
 *    🧮 Update Formulas - should refresh formulas on all sheets that have named ranges
 *    ✅ Update Outcomes Sheet Validations - tool to fix any issues with the data validation fields
 *    🔽 Deploy Extra Tracking Sheets - helps create all additional sheets for weekly pick 'em tracking
 *  🧙 AUTOMATION:
 *    📡 Spread Auto-Fetch Panel - lets you set a time for the schedule data (and spreads) to automatically be udpated
 *    ✅ Enable 👑&💀 Trigger - required for processing updates to Survivor/Eliminator evals (only visible if Survivor or Eliminator Present)
 *    ⭕ Disable 👑&💀 Trigger - remove if causing issues or want to run without it for a while (only visible if Survivor or Eliminator Present)
 * 
 *   ------------
 * 
 *   ❔ Help & Support - opens an HTML pop-up that has a link to send me an email and this project hosted on GitHub
 * 
 * ------------------------------------------------------------
 * 
 * If you're feeling generous and would like to support my work,
 * here's a link to support my wife, six kiddos, and me:
 * https://www.buymeacoffee.com/benpowers or https://venmo.com/benpowerscreative
 * 
 * Thanks for checking out the script!
 * 
 * **/
 
 /**
 * Runs when the spreadsheet is opened. Checks if the script has been
 * initialized for this document and either shows the authorization card or the main menu.
 */
function onOpen() {
  // Check a specific property to see if the first-run initialization is complete.
  const docProps = PropertiesService.getDocumentProperties();
  const isInitialized = docProps.getProperty('init') === 'true';
  if (isInitialized) {
    const config = JSON.parse(docProps.getProperty('configuration'));
    let contest = false;
    if (config) {
      if (config.survivorInclude || config.eliminatorInclude) {
        contest = true;
        survElimIcons = config.survivorInclude && config.eliminatorInclude ? '👑&💀' : config.survivorInclude ? '👑' : config.eliminatorInclude ? '💀' : false;
        survElimString = config.survivorInclude && config.eliminatorInclude ? 'Survivor/Eliminator' : config.survivorInclude ? 'Survivor' : config.eliminatorInclude ? 'Eliminator' : '';
      }
    }
    const ui = SpreadsheetApp.getUi();
    let menu = ui.createMenu('🏈 Picks')
    if (docProps.getProperty('configuration')) {
      if (contest) menu.addItem(`${survElimIcons} ${survElimString} Manager`,'launchSurvElimPanel');
      menu.addItem('📝 Form Builder', 'launchFormBuilder')
        .addItem('📋 Form Manager', 'launchFormManager')
        .addItem('📥 Check & Import Responses', 'launchFormImport');
    }
    menu.addItem('⚙️ Configuration', 'launchConfiguration')
      .addItem('👥 Member Manager', 'launchMemberPanel');
    menu.addSubMenu(ui.createMenu('✍️ Sign-Ups')
      .addItem('📝 Create or Open Sign-Up Form','launchSignupForm')
      .addItem('📥 Import Sign-Ups','importSignups'));
    
    if (docProps.getProperty('forms')) {
      menu.addSeparator()
      .addItem('✅ Close Out Week','closeOutWeek')
      .addItem(`🏈 Fetch ${LEAGUE} Outcomes`,'launchApiOutcomeImport')
      .addItem('📧 Remind Non-Respondents','remindNonRespondents')
      menu.addSubMenu(ui.createMenu('🧰 Utilities')
        .addItem(`📅 Update ${LEAGUE} Data`, 'fetchSchedule')
        .addItem('📊 Update Spread Data','fetchLatestSpreadsForWeek')
        .addItem('✏️ Rename a Member','showRenamePanel')
        .addItem('🧮 Update Formulas', 'allFormulasUpdate')
        .addItem('✅ Update Outcomes Sheet Validation','outcomesSheetUpdatePrompt')
        .addItem('🔽 Deploy Extra Tracking Sheets','setupSheets')
        .addItem('💵 Update Payouts','updatePayouts')
        .addItem('🩺 Health Check','healthCheck')
        .addItem('📸 Snapshot Spreadsheet','snapshotSpreadsheet'));
      let subMenu = ui.createMenu('🧙 Automation')
        .addItem('📡 Spread Auto-Fetch Panel','showAutoFetchPanel')
        .addItem('🔒 Kickoff Lock On/Off','toggleKickoffLock')
        .addItem('📧 Auto-Remind On/Off','toggleAutoRemind');
      if (contest) {
        subMenu.addItem(`✅ Enable ${survElimIcons} Triggers`,'createOnEditTrigger')
          .addItem(`⭕ Disable ${survElimIcons} Triggers`,'deleteOnEditTrigger');
      }
      menu.addSubMenu(subMenu);
    } else {
      menu.addSubMenu(ui.createMenu('🧰 Utilities')
        .addItem(`📅 Update ${LEAGUE} Data`, 'fetchSchedule')
        .addItem('📊 Update Spread Data','fetchLatestSpreadsForWeek'));
    }
    menu.addSeparator()
      .addItem('❔ Help & Support','showSupportDialog')
      .addToUi();

  } else {
    const ui = fetchUi();
    if (!docProps.getProperty('tz')){
      ui.alert(`🤗 WELCOME!`,`Thanks for checking out this script-enabled\nsheet for running any combination of these pools:\n\n\u2003\u2003🔹 Pick 'Ems (🏈)\n\u2003\u2003🔹 Survivor (👑)\n\u2003\u2003🔹 Eliminator (💀)\n\nBefore you get started, you'll need to allow the scripts\nto run and ensure your time zone is set correctly.`, ui.ButtonSet.OK);
      timezoneCheck(ui,docProps);
      ui.alert(`⏩ Next`,`Now run the "🟢 Initialize" script from\nthe "🏈 Picks" menu along the top bar.`, ui.ButtonSet.OK);
    } else {
      ui.alert(`🤗 WELCOME!`,`Thanks for checking out this script-enabled\nsheet for running any combination of these pools:\n\n\u2003\u2003🔹 Pick 'Ems (🏈)\n\u2003\u2003🔹 Survivor (👑)\n\u2003\u2003🔹 Eliminator (💀)\n\nBefore you get started, you'll need to allow the scripts\nto run if not already authorized and initialize the sheet.\n\nRun "🟢 Initialize" from the "🏈 Picks" menu along the top bar`, ui.ButtonSet.OK);
    }
    SpreadsheetApp.getUi()
      .createMenu('🏈 Picks')
      .addItem('🟢 Initialize', 'launchConfiguration')
      .addToUi();
  }
}

//------------------------------------------------------------------------
// CHECKS IF TZ PROP SET AND PROMPTS IF NOT
function timezoneCheck(ui,docProps) {
  if (docProps == undefined) {
    docProps = PropertiesService.getDocumentProperties();
  }
  ui = fetchUi(ui);
  const tzProp = docProps.getProperty('tz');
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  // Confirm timezone setting before continuing
  if (tzProp == null) {
    let timeZonePrompt = ui.alert(`🕐 TIMEZONE`,`The timezone you're currently using is ${tz}.\n\nIs this correct?`, ui.ButtonSet.YES_NO);
    if ( timeZonePrompt != 'YES') {
      ui.alert(`🩹 FIX TIMEZONE`, `Follow these steps to change your projects time zone:\n\n1. Go to the "Extensions" > "Apps Script" menu\n2. Select the gear icon on the left menu\n3. Use the drop-down to select the correct timezone\n4. Close the "Apps Script" editor and return to the sheet\n5. Restart the script through the "🏈 Picks" menu`, ui.ButtonSet.OK);
      return false;
    } else if ( timeZonePrompt == 'YES') {
      docProps.setProperty('tz',tz);
      return true;
    }
  } else {
    return true;
  }
}

// ============================================================================================================================================
// GLOBAL VARIABLES
// ============================================================================================================================================

// GLOBAL VARIABLES - for easy modification in the future
const LEAGUE = "NFL"; // Hopefully I"ll be able to support NCAAF at some point
const TEAMS = 32;
const REGULAR_SEASON = 18; // Regular season matchups
const WEEKS = 23; // Total season weeks (including playoffs)
const WEEKS_TO_EXCLUDE = [22]; // Break before Superbowl
const MAXGAMES = TEAMS/2;
const SCOREBOARD = 
    LEAGUE == "NFL" ? "https://site.web.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" :
    (LEAGUE == "NCAAF" ? "https://site.web.api.espn.com/apis/site/v2/sports/football/college-football/scoreboard" : null);
const COLOR_PRIMARY = "#D50A0A";
const COLOR_SECONDARY = "#D50A0A";
const COLOR_TERTIARY = "#FFFFFF";

const DAY = {  0: {"name":"Sunday","index":0},  1: {"name":"Monday","index":1},  2: {"name":"Tuesday","index":2},  3: {"name":"Wednesday","index":-4},  4: {"name":"Thursday","index":-3},  5: {"name":"Friday","index":-2},  6: {"name":"Saturday","index":-1} };

const WEEKNAME = { 19: {"name":"WildCard","teams":12,"matchups":6}, 20: {"name":"Divisional","teams":8,"matchups":4}, 21: {"name":"Conference","teams":4,"matchups":2}, 23: {"name":"SuperBowl","teams":2,"matchups":1} };

const weeklySheetPrefix = "WK";
const schedulePrefix = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/";
const scheduleSuffix = "?view=proTeamSchedules";

// ESPN fetch hardening (see espnFetch in UTILITIES) — Google's servers get 403'd without these
const ESPN_USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36";
const ESPN_FETCH_ATTEMPTS = 3;
const ESPN_FETCH_BACKOFF_MS = 800;
const fallbackYear = 2026;
const pickColors = {
  correct:   { even: '#c9ffdf', odd: '#a0fdba', end: '#69ffa6', font: '#000000' },
  incorrect: { even: '#fff7f9', odd: '#fff2f4', end: '#fcd4dc', font: '#999999' },
  home:      { even: '#e3fffe', odd: '#d0f5f3', end: '#80f1ea', font: '#000000' },
  away:      { even: '#fffee3', odd: '#faf9e1', end: '#fbf77f', font: '#000000' }
};
const dayColorsObj = {"Wednesday": "#ffe3cc","Thursday":"#fffdcc","Friday":"#e7fed1","Saturday":"#cffdda","Sunday":"#bbfbe7","Monday":"#adf7f5"};
const dayColorsFilledObj = {"Wednesday": "#ffeb90","Thursday":"#fffb95","Friday":"#d4ffa6","Saturday":"#abffbf","Sunday":"#89fddb","Monday":"#74f7f3"};
const dayColors = ["#ffe3cc", "#fffdcc","#e7fed1","#cffdda","#bbfbe7","#adf7f5"];
const dayColorsFilled = ["#ffeb90", "#fffb95","#d4ffa6","#abffbf","#89fddb","#74f7f3"];
const configTabColor = "#ff9561";
const leaderboardTabColor = "#55ff90";

// ============================================================================================================================================
// POOL & PAYOUT CONFIGURATION
// ============================================================================================================================================
//
// >>> SET THESE FEES TO YOUR OWN POOL'S NUMBERS BEFORE RUNNING A SEASON. <<<
//
// Every fee below is a $1 placeholder so the arithmetic is easy to follow. They are not
// meant to be used as-is.
//
// "fee" is charged per entry. "weeks" multiplies the per-entry cost, so a weekly pool's pot
// for a SINGLE week is fee x entries, while its season-long contribution is
// fee x entries x weeks. Every pot in the workbook is derived from these two numbers plus
// the entry count, so changing a fee here and re-running "Update Payouts" flows it through
// every tab -- there are no hardcoded dollar amounts anywhere else.
//
// With these placeholders the entry total is $44 per member
// (18 + 18 + 1 + 1 + 1 + 4 + 1), and at 25 entries the weekly pot would be $25.
//
// Drop a pool you do not run by deleting its line. Its column disappears from
// Summary Payout and its pot stops being collected.
const POOLS = [
  { key: 'weeklyWins',      label: 'Weekly Wins',                    fee: 1,  weeks: 18, cadence: 'weekly'   },
  { key: 'weeklyConsensus', label: 'Weekly Consensus',               fee: 1,  weeks: 18, cadence: 'weekly'   },
  { key: 'seasonPct',       label: 'Season % Correct (best 16 wks)', fee: 1,  weeks: 1,  cadence: 'season'   },
  { key: 'seasonPoints',    label: 'Season Total Points',            fee: 1,  weeks: 1,  cadence: 'season'   },
  { key: 'seasonConsensus', label: 'Season Consensus',               fee: 1,  weeks: 1,  cadence: 'season'   },
  { key: 'postSeason',      label: 'Post Season',                    fee: 1,  weeks: 4,  cadence: 'playoffs' },
  { key: 'perfectWeek',     label: 'Perfect Week Pool',              fee: 1,  weeks: 1,  cadence: 'season'   }
];

// Reminders to anyone who has not submitted go out this many hours before the week's first
// kickoff. The first kickoff is used rather than "Thursday" specifically, because some weeks
// open Wednesday (2026 week 1) or have Friday games -- this always lands before any game starts.
const REMINDER_HOURS_BEFORE_KICKOFF = 4;

// Pool 3 keeps each player's best 16 weekly percentages out of the 18 regular-season weeks.
const SEASON_PCT_WEEKS = 16;

// The playoff weeks that the Post Season pool scores. Week 22 is the pre-Super Bowl bye and
// is already excluded by WEEKS_TO_EXCLUDE; the Pro Bowl is never scored.
const PLAYOFF_WEEKS = [19, 20, 21, 23];

// Percentage of a pot paid to each finishing place. The number of paid places scales with
// the size of the pool -- set before the season and left alone once it starts.
const PAYOUT_LADDERS = {
  3: [50, 30, 20],
  4: [45, 27, 18, 10],
  5: [40, 25, 20, 10, 5],
  7: [34, 22, 16, 11, 8, 5, 4]
};

// Entry-count breakpoints deciding how many places get paid.
const PAYOUT_PLACE_BREAKPOINTS = [
  { minEntries: 35, places: 7 },
  { minEntries: 25, places: 5 },
  { minEntries: 15, places: 4 },
  { minEntries: 0,  places: 3 }
];

const generalTabColor = "#aaaaaa";
const winnersTabColor = "#ffee00";
const survElimTabColors = {"survivor":"#ffee00","eliminator":"#fca503"}
const nameValidation = FormApp.createTextValidation()
  .setHelpText('Enter a minimum of 2 characters, up to 30.')
  .requireTextMatchesPattern(".{2,30}")
  .build();
  
const scheduleTabColor = "#472a24";
const numberMap = {10:"🔟", 9:"9️⃣", 8:"8️⃣", 7:"7️⃣", 6:"6️⃣", 5:"5️⃣", 4:"4️⃣", 3:"3️⃣", 2:"2️⃣", 1:"1️⃣", 0:"0️⃣" };  

const LEAGUE_DATA = {
  "ARI": {
    "division": "NFC West",
    "division_opponents": ["LAR", "SEA", "SF"],
    "colors": [
      "#97233F",
      "#000000",
      "#FFB612",
      "#FFFFFF"
    ],
    "mascot": "🐦",
    "colors_emoji": "🔴⚫"
  },
  "ATL": {
    "division": "NFC South",
    "division_opponents": ["CAR", "NO", "TB"],
    "colors": [
      "#010101",
      "#A6192E",
      "#FFFFFF",
      "#B2B4B2"
    ],
    "mascot": "🦜",
    "colors_emoji": "⚫🔴"
  },
  "BAL": {
    "division": "AFC North",
    "division_opponents": ["CIN", "CLE", "PIT"],
    "colors": [
      "#24125F",
      "#FFFFFF",
      "#9A7611",
      "#010101"
    ],
    "mascot": "🐦‍⬛",
    "colors_emoji": "🟣🟡"
  },
  "BUF": {
    "division": "AFC East",
    "division_opponents": ["MIA", "NE", "NYJ"],
    "colors": [
      "#003087",
      "#C8102E",
      "#FFFFFF",
      "#091F2C"
    ],
    "mascot": "🐃",
    "colors_emoji": "🔵🔴"
  },
  "CAR": {
    "division": "NFC South",
    "division_opponents": ["ATL", "NO", "TB"],
    "colors": [
      "#101820",
      "#0085CA",
      "#B2B4B2",
      "#FFFFFF"
    ],
    "mascot": "🐈‍⬛",
    "colors_emoji": "⚫🔵"
  },
  "CHI": {
    "division": "NFC North",
    "division_opponents": ["DET", "GB", "MIN"],
    "colors": [
      "#091F2C",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐻",
    "colors_emoji": "🔵🟠"
  },
  "CIN": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CLE", "PIT"],
    "colors": [
      "#010101",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐅",
    "colors_emoji": "⚫🟠"
  },
  "CLE": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "PIT"],
    "colors": [
      "#311D00",
      "#EB3300",
      "#FFFFFF",
      "#EDC8A3"
    ],
    "mascot": "🟠",
    "colors_emoji": "🟤🟠"
  },
  "DAL": {
    "division": "NFC East",
    "division_opponents": ["NYG", "PHI", "WSH"],
    "colors": [
      "#0C2340",
      "#FFFFFF",
      "#87909A",
      "#7F9695"
    ],
    "mascot": "🤠",
    "colors_emoji": "🔵⚪"
  },
  "DEN": {
    "division": "AFC West",
    "division_opponents": ["KC", "LAC", "LV"],
    "colors": [
      "#0C2340",
      "#FC4C02",
      "#FFFFFF"
    ],
    "mascot": "🐴",
    "colors_emoji": "🔵🟠"
  },
  "DET": {
    "division": "NFC North",
    "division_opponents": ["CHI", "GB", "MIN"],
    "colors": [
      "#0069B1",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "🦁",
    "colors_emoji": "⚪🔵"
  },
  "GB": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "MIN"],
    "colors": [
      "#183029",
      "#FFB81C",
      "#FFFFFF"
    ],
    "mascot": "🧀",
    "colors_emoji": "🟢🟡"
  },
  "HOU": {
    "division": "AFC South",
    "division_opponents": ["IND", "JAX", "TEN"],
    "colors": [
      "#1D1F2A",
      "#E4002B",
      "#FFFFFF",
      "#0072CE"
    ],
    "mascot": "🐂",
    "colors_emoji": "🔴🔵"
  },
  "IND": {
    "division": "AFC South",
    "division_opponents": ["HOU", "JAX", "TEN"],
    "colors": [
      "#003A70",
      "#FFFFFF",
      "#A2AAAD",
      "#1D252D"
    ],
    "mascot": "🐎",
    "colors_emoji": "🔵⚪"
  },
  "JAX": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "TEN"],
    "colors": [
      "#006271",
      "#D29F13",
      "#010101",
      "#9A7611"
    ],
    "mascot": "🌴",
    "colors_emoji": "🟡🔵"
  },
  "KC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "LAC", "LV"],
    "colors": [
      "#C8102E",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🏹",
    "colors_emoji": "🔴🟡"
  },
  "LAC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LV"],
    "colors": [
      "#0072CE",
      "#FFB81C",
      "#FFFFFF",
      "#0C2340"
    ],
    "mascot": "⚡",
    "colors_emoji": "🔵🟡"
  },
  "LAR": {
    "division": "NFC West",
    "division_opponents": ["ARI", "SEA", "SF"],
    "colors": [
      "#1E22AA",
      "#FFD100",
      "#D7D2CB",
      "#FFFFFF"
    ],
    "mascot": "🐏",
    "colors_emoji": "🔵🟡"
  },
  "LV": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LAC"],
    "colors": [
      "#010101",
      "#A2AAAD",
      "#FFFFFF",
      "#87909A"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "⚫⚪"
  },
  "MIA": {
    "division": "AFC East",
    "division_opponents": ["BUF", "NE", "NYJ"],
    "colors": [
      "#008C95",
      "#FC4C02",
      "#FFFFFF",
      "#005776"
    ],
    "mascot": "🐬",
    "colors_emoji": "🟡🔵"
  },
  "MIN": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "GB"],
    "colors": [
      "#582C83",
      "#FFC72C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⚔️",
    "colors_emoji": "🟣🟡"
  },
  "NE": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NYJ"],
    "colors": [
      "#0C2340",
      "#C8102E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🥁", //🧦
    "colors_emoji": "🔵🔴"
  },
  "NO": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "TB"],
    "colors": [
      "#010101",
      "#D3BC8D",
      "#FFFFFF",
      "#A28D5B"
    ],
    "mascot": "⚜️",
    "colors_emoji": "⚫🟡"
  },
  "NYG": {
    "division": "NFC East",
    "division_opponents": ["DAL", "PHI", "WSH"],
    "colors": [
      "#001E62",
      "#A6192E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🏗️",
    "colors_emoji": "🔵🔴"
  },
  "NYJ": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NE"],
    "colors": [
      "#115740",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "✈️",
    "colors_emoji": "🟢⚪"
  },
  "PHI": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "WSH"],
    "colors": [
      "#004851",
      "#E3E5E6",
      "#545859",
      "#010101"
    ],
    "mascot": "🦅",
    "colors_emoji": "🟢⚫"
  },
  "PIT": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "CLE"],
    "colors": [
      "#010101",
      "#FFB81C",
      "#FFFFFF",
      "#C8102E"
    ],
    "mascot": "🏭",
    "colors_emoji": "⚫🟡"
  },
  "SEA": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SF"],
    "colors": [
      "#0C2340",
      "#78BE21",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🌊",
    "colors_emoji": "🔵🟢"
  },
  "SF": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SEA"],
    "colors": [
      "#A6192E",
      "#B9975B",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⛏️",
    "colors_emoji": "🔴🟡"
  },
  "TB": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "NO"],
    "colors": [
      "#010101",
      "#A6192E",
      "#3D3935",
      "#DC4405"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "🔴⚫"
  },
  "TEN": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "JAX"],
    "colors": [
      "#0C2340",
      "#418FDE",
      "#B2B4B2",
      "#C8102E"
    ],
    "mascot": "🛡️",
    "colors_emoji": "🔵🔴"
  },
  "WSH": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "PHI"],
    "colors": [
      "#651C32",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🎖️",
    "colors_emoji": "🟤🟡"
  }
};

/**
 * Displays a custom HTML modal dialog to guide the user through authorization.
 */
function showAuthorizationCard() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Montserrat', Arial, sans-serif; padding: 10px; text-align: center; }
          h2 { color: #013369; margin-top: 0px; }
          p { font-size: 14px; color: #333; line-height: 1.6; }
          .permissions { text-align: left; background-color: #f8f9fa; border: 1px solid #e8eaed; border-radius: 8px; padding: 12px; margin-top: 15px; }
          .permissions strong { color: #D50A0A; }
          .btn { background-color: #013369; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: 600; margin-top: 25px; }
          .btn:hover { background-color: #2067b3; }
        </style>
      </head>
      <body>
        <h2>Welcome to NFL Picks!</h2>
        <p>Before you can get started, the scripts needed your permission to run.</p>
        
        <div class="permissions">
          <strong>Why did you need to grant these permissions?</strong>
          <ul>
            <li><strong>View and manage Sheets:</strong> To create and update this Sheet.</li>
            <li><strong>View and manager Drive:</strong> To create and update your weekly picks Forms.</li>
            <li><strong>Connect to an external service:</strong> To fetch the latest ${LEAGUE} game schedules and scores.</li>
          </ul>
        </div>

        <p style="margin-bottom: 0;">Feel free to review the code within the "Extensions" > "Apps Script". You should be all set to create your pool now. Thanks for checking out this tool and please be patient as improvements and fixes are made.</p>

        <button class="btn" style="margin-top: 4px; padding: 10px 16px;" onclick="authorizeScript()">Let's Go!</button>

        <script>
          function authorizeScript() {
            google.script.run
              .withSuccessHandler(onAuthorizationSuccess)
              .withFailureHandler(onAuthorizationFailure)
              .triggerAuthorizationFlow();
          }
          
          function onAuthorizationSuccess() {
            // The success message is now a clear call to action.
            alert('Authorization successful! Please click "Configuration" from the menu again to open the config panel.');
            google.script.host.close();
          }

          function onAuthorizationFailure(err) {
            alert('Authorization failed or was canceled. Please try again. Error: ' + err.message);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'One-Time Setup Required');
}

/**
 * A simple server-side function whose only purpose is to trigger the auth flow.
 * Because this function requires a permission (accessing properties), calling it
 * will force the Google authorization dialog to appear if the user isn't yet authorized.
 */
function triggerAuthorizationFlow() {
  // This line requires permission, which is what we need.
  PropertiesService.getDocumentProperties().setProperty('init', 'true');
  
  // By setting this property, the next time onOpen runs, it will create the real menu.
  Logger.log(`✅ Script has been successfully initialized and authorized for this document.`);
  onOpen();
}

/**
 * Will check for authorization, if authorized will launch configuration panel
 */
function launchConfiguration() {
  const isInitialized = PropertiesService.getDocumentProperties().getProperty('init');
  
  if (isInitialized === 'true') {
    // If we're initialized, show the real sidebar.
    configurationPanel();
  } else {
    // If this is the first run, show the authorization walk-through.
    showAuthorizationCard();
  }
}

/**
 * Creates an HTML-based configuration sidebar that mimics CardService styling.
 * This is the recommended approach for Google Sheets add-ons.
 */
function configurationPanel() {
  const html = HtmlService.createHtmlOutputFromFile('configurationPanel.html')
    .setTitle(`${LEAGUE} Pool Configuration`) 
    .setWidth(350); 
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Processes the configuration from the sidebar.
 * Uses the saveProperties() helper function.
 */
function processConfigurationSubmission(formObject) {
  const docProps = PropertiesService.getDocumentProperties();
  const previousConfig = docProps.getProperty('configuration');
  try {
    const allOptions = [
      'groupName',

      'pickemsInclude', // Pick 'Ems options below
      'pickemsAts',
      'mnfExclude',
      'commentsExclude',
      'bonusInclude',
      'mnfDouble',
      'tiebreakerInclude',
      'overUnderInclude',
      'weeklyPaidTracking',
      'hideNonParticipants',
      
      'survivorInclude', // Survivor options below
      'survivorStartWeek',
      'survivorAts',
      'survivorRevives',
      'survivorLives',
      
      'eliminatorInclude', // Elminator options below
      'eliminatorStartWeek',
      'eliminatorAts',
      'eliminatorRevives',
      'eliminatorLives',
      
      'customizeMatchups', // Matchup customization
      'customizeMode',
      'matchupCustomization',

      'membershipLocked',
      'kickoffLock', // General options below
      // 'playoffsExclude',
      'hideEmojis',
      'initialized',
      'year'
      
    ];
    const configToSave = {};
    let removeNewUserEntry = false;
    allOptions.forEach(option => {
      if (option === 'membershipLocked') {
        if (previousConfig.hasOwnProperty('membershipLocked')) {
          if (formObject.hasOwnProperty('membershipLocked')) {
            if (previousConfig.membershipLocked === false && formObject.membershipLocked === true) {
              removeNewUserEntry = true;
            }
          }
        }
      }
      if (formObject.hasOwnProperty(option)) {
        configToSave[option] = formObject[option];
      };
    });

    if (!configToSave.groupName) {
      if (configToSave.pickemsInclude) {
        configToSave.groupName = `${LEAGUE} Pick 'Ems`;
      } else {
        configToSave.groupName = `${LEAGUE} Survivor Pool`;
      }
    }
  
    let modes = ['survivor','eliminator'];
    let week;
    for (const type in modes) {
      if (configToSave[`${modes[type]}Include`]) {
        week = week || fetchWeek();
        if (parseInt(previousConfig[`${modes[type]}StartWeek`]) < parseInt(configToSave[`${modes[type]}StartWeek`])) {
          if (parseInt(configToSave[modes[type]+'StartWeek']) < week) {
            SpreadsheetApp.getUi().alert(`⚠️ ${type.toUpperCase()} START WEEK ISSUE!`, `You set your ${type} pool to week ${configToSave[modes[type]+'StartWeek']}, which is prior to this week (${week}). Please update to restart the pool`,SpreadsheetApp.getUi().ButtonSet.OK)
            configToSave[`${modes[type]}Active`] = false;
          }
          if (parseInt(configToSave[modes[type]+'StartWeek']) == week ) {
            Logger.log(`✅ New start week for ${type} pool is in the future, will start in week ${configToSave[modes[type]+'StartWeek']} (formerly started in week ${previousConfig[`${modes[type]}StartWeek`]})`);
          } else {
            Logger.log(`⚠️ New start week for ${type} pool is this week; ensure new form is generated with this change if one exists! (formerly started in week ${previousConfig[`${modes[type]}StartWeek`]})`);
          }
          configToSave[`${modes[type]}Active`] = true;
        } else if (!configToSave[`${modes[type]}Active`]) {
          configToSave[`${modes[type]}Active`] = false;
        }
      }
      if ( configToSave[`${modes[type]}Include`] && !previousConfig[`${modes[type]}Include`] ) {
        configToSave[`${modes[type]}Active`] = true;
      }
    }

    if (!configToSave.year) configToSave.year = fetchYear();
    if (!configToSave.hasOwnProperty('initialized')) configToSave.initialized = false;
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${LEAGUE}_OUTCOMES`)) {
      updateOutcomeSheetVisibility(configToSave);
    }
    saveProperties('configuration', configToSave);

    if (removeNewUserEntry) {
      try {
        const forms = JSON.parse(docProps.getProperty('forms'));
        if (!forms) {
          Logger.log(`⭕ No forms exist for removing a new user from`);
        } else {
          let maxWeek = Math.max(...Object.keys(data).map(key => parseInt(key)));
          removeNewUserQuestion(maxWeek);
          Logger.log(`❌ Removed new user from most recent form`);
        }
      }
      catch (err) {
        Logger.log(`Error removing "New User" question from form or it wasn't present: ${err.stack}`);
      }
    }
    // Recreate menu...
    onOpen();
    
    return { success: true, data: configToSave };

  } catch (err) {
    Logger.log(`Error in processConfigurationSubmission: ${err.stack}`);
    throw new Error(`❗ Failed to create configuration: ${err.toString()}`);
  }
}

/**
 * Loads the parse JSON object from Document Properties.
 * If no object is found, it returns an empty object.
 *
 * @returns {Object} The parsed object.
 */
function fetchProperties(name) {
  const string = PropertiesService.getDocumentProperties().getProperty(name);
  if (string) {
    return JSON.parse(string, (key, value) => {
      if (typeof value == 'string') {
        if (value === "true") {
          return true;
        } else if (value === "false") {
          return false;
        }
      }
      return value; // Return the value as is if not "true" or "false" string
      });
  } else {
    // If no property are found, return a default state.
    return {};
  }
}

/**
 * Saves a given JavaScript object to Document Properties as a single JSON string.
 *
 * @param {Object} object The configuration object to save.
 */
function saveProperties(name,object) {
  if (!object || typeof object !== 'object' || Object.keys(object).length === 0) {
    Logger.log(`❗ saveProperties was called with an invalid or empty object. No properties were set.`);
    return; // Exit the function if there's nothing to set.
  }
  const propString = JSON.stringify(object);
  PropertiesService.getDocumentProperties().setProperty(name, propString);
}

/**
 * Deletes a single Document Property
 * If no object is found, it returns an empty object.
*/
function deleteProperties(name) {
  try {
    PropertiesService.getDocumentProperties().deleteProperty(name);
    return true;
  } catch (err) {
    Logger.log(`Error removing property ${name}: ${err.stack}`);
    return false;
  }
}


/**
 * NOT NECESSARY for final deployment
 * Remove configuration data from Document Properties
 */
function deleteConfiguration() {
  try {
    deleteProperties('configuration');
    SpreadsheetApp.getActiveSpreadsheet().toast(`SUCCESS: "configuration" removed`);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`FAILURE: "configuration" not removed, error: ${err.stack}`);
  }
}

/**
 * Quick call to fetch via fetchProperties('configuration') with a check for existing and then providing an empty config if not found
 */
function fetchConfiguration(provideTemplate) {
  try {
    const props = fetchProperties('configuration');
    if (Object.keys(props).length > 0) {
      if (props.year) {
        if (!props.pickemsInclude) {
          props.pickemsInclude = true;
          props.tiebreakerInclude = true;
        }
        return props;
      } else {
        try {
          props.year = fetchYear(true);
        } catch (err) {
          Logger.log(`Issue fetching year when the configuration already existed within fetchConfiguration() call: ${err.stack}`);
          props.year = '2025';
        }
        return props;
      }
    } else {
      Logger.log(`⭕ No existing configuration sidebar data, presenting template object`);
    }
  } catch (err) {
    Logger.log(`❌ Failed to retrieve configuration sidebar data: ${err.stack}`);
  }
  let placeholder = {pickemsInclude:true,tiebreakerInclude:true};
  if (provideTemplate) {
    try {
      placeholder.year = fetchYear(true);
    } catch (err) {
      Logger.log(`⚠️ Issue fetching year during the placeholder configuration creation within fetchConfiguration(provideTemplate) call: ${err.stack}`);
      placeholder.year = '2025';
    }
    Logger.log(`↩️ Returning placeholder: ${JSON.stringify(placeholder)}`);
    return placeholder;
  } else {
    SpreadsheetApp.getUi().alert('No configuration found, starting configuration process...');
    showConfigDialog();
  }
}

/**
 * Retrieves all necessary data for the configuration sidebar on load.
 * Includes a calculated message about Sunday kickoff times.
 */
function fetchConfigurationSidebarData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const config = fetchConfiguration(true); // Includes Pick 'Ems ON and YEAR
    config.week = fetchWeek();
    const scriptTimeZone = Session.getScriptTimeZone();
    const formattedTime = Utilities.formatDate(new Date(), scriptTimeZone, "h:mm a',' EEE, MMM d");
    // Calculate the local kickoff time
    let kickoffMessage = '';
    try {
      // Create a reference Date object for 1:00 PM in New York (Eastern Time).
      // We use Utilities.parseDate for a reliable way to create a date in a specific timezone.
      // The date itself doesn't matter, only the time and zone.
      const kickoffTimeET = Utilities.parseDate("13:00", "America/New_York", "HH:mm");

      // Format that same moment in time for the user's detected scriptTimeZone.
      // 'h a' will format as "10 AM", "1 PM", etc.
      const localKickoffTime = Utilities.formatDate(kickoffTimeET, scriptTimeZone, "h a");

      // Construct the helpful message.
      kickoffMessage = `Sunday ${LEAGUE} games will show a ${localKickoffTime} kickoff.`;

    } catch (err) {
      Logger.log(`❌ Could not calculate local kickoff time. ${err.stack}`);
      kickoffMessage = `❌ Could not calculate local kickoff time.`;
    }
    const obj = {
      properties: config,
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: {
        zone: scriptTimeZone,
        currentTime: formattedTime,
        kickoffMessage: kickoffMessage
      }
    };
    return {
      properties: config,
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: {
        zone: scriptTimeZone,
        currentTime: formattedTime,
        kickoffMessage: kickoffMessage
      }
    };
  } catch (err) {
    Logger.log(`⚠️ Error preparing config sidebar data: ${err.stack}`);
    return {
      properties: {pickemsInclude: true},
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: { zone: 'Unknown', currentTime: 'N/A', kickoffMessage: '' }
    };
  }
}

//------------------------------------------------------------------------
// SUPPORT POPUP FOR HELP - Loads HTML "supportPrompt.html" file
function showSupportDialog() {
  let html = HtmlService.createHtmlOutputFromFile('supportPrompt.html')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// Returns version to help and support popup to quickly review version by users
function getSupportPromptInfo() {
  Logger.log(`↩️ Returning support prompt info, version: ${VERSION}`);
  return {
    version: VERSION
  };
}

//------------------------------------------------------------------------
// CONTINUATION OF SETUP - After a successful submission of the HTML prompt, this script picks up for some finishing questions and then runs the setup
function setupSheets() {
  const ss = fetchSpreadsheet();
  const docProps = PropertiesService.getDocumentProperties();
  let config = JSON.parse(docProps.getProperty('configuration'));
  
  if (!config) {
    launchConfiguration();
    ss.toast('Configuration not found or not set up yet, launching now...','⚠️ CONFIGURATION NEEDED');
    return
  }
  const memberData = JSON.parse(docProps.getProperty('members'));
  if (!memberData) {
    launchMemberPanel();
    ss.toast('Members not found or not set up yet, launching now...','⚠️ MEMBERS NEEDED');
    return;
  }
  // This rebuilds tabs in place, so offer a full copy first once there is data worth keeping
  if (config.initialized) {
    const ui = fetchUi();
    const backup = ui.alert(`📸 SNAPSHOT FIRST?`,
      `Deploying the tracking sheets rebuilds them in place.\n\nSave a timestamped copy of this spreadsheet first?`,
      ui.ButtonSet.YES_NO_CANCEL);
    if (backup === ui.Button.CANCEL) return;
    if (backup === ui.Button.YES) snapshotSpreadsheet(true);
  }
  try {
    const year = fetchYear();
    let week = fetchWeek();
    
    if (!ss.getSheetByName(`${LEAGUE}_OUTCOMES`)) outcomesSheet(ss);

    Logger.log(`Deployed ${LEAGUE} Outcomes sheet`);
    if (config.pickemsInclude) {
      // Creates Weekly Totals Record Sheet
      totSheet(ss,memberData);
      Logger.log('Deployed Weekly Totals sheet');
      ss.toast('Deployed Weekly Totals sheet');

      // Creates Weekly Rank Record Sheet
      rnkSheet(ss,memberData);
      Logger.log('Deployed Weekly Rank sheet');
      ss.toast('Deployed Weekly Rank sheet');
      
      // Creates Weekly Percentages Record Sheet
      pctSheet(ss,memberData);
      Logger.log('Deployed Weekly Percentages sheet');
      ss.toast('Deployed Weekly Percentages sheet');

      // Creates the money tabs (Wkly Payout, Wkly Consensus, Season Points, Summary Payout)
      updatePayoutSheetsQuietly(ss);
      Logger.log('Deployed payout sheets');
      ss.toast('Deployed payout sheets');
    
      // Creates Winners Sheet
      winnersSheet(ss,year);
      Logger.log('Deployed Winners sheet');
      ss.toast('Deployed Winners sheet');
      
      // Creates MNF Sheet
      if (!config.mnfExclude) {
        mnfSheet(ss,memberData);
        Logger.log('Deployed MNF sheet');
        ss.toast('Deployed MNF sheet');
      }
    }
    if (config.survivorInclude) {
      // Creates Survivor Sheet
      let survivor = survElimSheet(ss,config,memberData,'survivor');
      
      Logger.log('Deployed Survivor sheet');
      ss.toast('Deployed Survivor sheet');

      if (!config.pickemsInclude) {
        survivor.activate();
      }
    } else {
      try{ss.deleteSheet(ss.getSheetByName('SURVIVOR'));} catch (err) {}
    }

    if (config.eliminatorInclude) {
      // Creates Eliminator Sheet
      let eliminator = survElimSheet(ss,config,memberData,'eliminator');
      
      Logger.log('Deployed Eliminator sheet');
      ss.toast('Deployed Eliminator sheet');

      if (!config.pickemsInclude) {
        eliminator.activate();
      }
    } else {
      try{ss.deleteSheet(ss.getSheetByName('ELIMINATOR'));} catch (err) {}
    }     
    
    // Creates Summary Record Sheet
    summarySheet(ss,memberData,config);
    Logger.log('Deployed Summary sheet');

    ss.getSheetByName(LEAGUE).hideSheet();

    let sheet = ss.getSheetByName('Sheet1');
    if ( sheet != null ) {
      ss.deleteSheet(sheet);
    }
    Logger.log(`Deleted 'Sheet 1'`);
    
    Logger.log(`You're all set, have fun!`);
    config.initialized = true;
    saveProperties('configuration', config);
  }
  catch (err) {
    Logger.log(`runFirstStack ${err.stack}`);
  }
}

// ============================================================================================================================================
// MEMBER LIST FUNCTIONS
// ============================================================================================================================================

/**
 * Processes a revive request for a specific member and game type.
 * It fetches the current member data, updates the relevant properties for the
 * specified member, and saves the data back.
 *
 * @param {Object} data An object from the client, e.g., { memberId: "id_123...", gameType: "survivor" }.
 * @returns {Object} A success or error message object to be sent back to the client.
 */
function processReviveMember(data) {
  const { memberId, gameType, week } = data;

  if (!memberId || !gameType || !week) {
    throw new Error("Invalid request. Missing ID, Type, or Week.");
  }

  try {
    const memberData = fetchProperties('members');
    const member = memberData.members[memberId];
    const config = fetchProperties('configuration');

    const livesKey = gameType === 'survivor' ? 'sL' : 'eL';
    const revivesKey = gameType === 'survivor' ? 'sR' : 'eR';
    const eliminatedKey = `${gameType.substring(0,1).toLowerCase()}O`;
    const startingLives = parseInt(config[`${gameType}Lives`], 10) || 1;
    
    if (!member[livesKey]) member[livesKey] = [];

    // --- 1. LIVES RESET ---
    member[livesKey][week - 1] = startingLives;
    // Clear out subsequent zeros to ensure they stay revived in future weeks
    for (let i = week; i < member[livesKey].length; i++) {
      member[livesKey][i] = startingLives;
    }

    // --- 2. REVIVE ARRAY LOGIC ---
    // Ensure the revive property is an array
    if (!Array.isArray(member[revivesKey])) {
      let legacyValue = parseInt(member[revivesKey]) || 0;
      member[revivesKey] = [];
      // If there was an old integer count, we put it at index 0 (Week 1) as a placeholder
      if (legacyValue > 0) member[revivesKey][0] = legacyValue;
    }
    
    // Increment the count for the specific week index
    const currentVal = member[revivesKey][week - 1] || 0;
    member[revivesKey][week - 1] = currentVal + 1;
    
    // --- 3. STATUS CLEANUP ---
    delete member[eliminatedKey];
    
    saveProperties('members', memberData);

    return { 
      success: true, 
      message: `🌅 ${member.name} has been revived for Week ${week}!`,
      updatedMemberData: memberData 
    };

  } catch (err) {
    Logger.log(`⚠️ Error in "processReviveMember": ${err.stack}`);
    throw new Error(`⚠️ Failed to process revive: ${err.message}`);
  }
}

/**
 * Retrieves all data needed for the Member Management panel.
 * This now includes the member list AND a boolean indicating if the
 * 'Revive' feature should be displayed, based on the main pool configuration.
 *
 * @returns {Object} An object containing memberData and a showReviveButtons flag.
 */
function getMembersSidebarData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();

    // Fetch the member data
    const members = JSON.parse(docProps.getProperty('members')) || { order: [], details: {} }; // Ensure a default object

    // Fetch the main configuration
    const config = JSON.parse(docProps.getProperty('configuration')) || {}; // Ensure a default object

    // Calculate the new boolean flag based on the required conditions
    // This will be false if either setting is false or doesn't exist
    const showReviveSurvivorButtons = (config.survivorInclude === true && config.survivorRevives === true);
    const showReviveEliminatorButtons = (config.eliminatorInclude === true && config.eliminatorRevives === true);

    // Return a single, bundled object with all the data the client needs
    return {
      week: fetchWeek() || 1,
      memberData: members,
      showReviveSurvivorButtons: showReviveSurvivorButtons,
      showReviveEliminatorButtons: showReviveEliminatorButtons
    };

  } catch (err) {
    Logger.log(`⚠️ Error preparing member panel data: ${err.stack}`);
    // Return a safe, default structure in case of any error.
    return {
      week: fetchWeek() || 1,
      memberData: { membersOrder: [], members: {} },
      showReviveSurvivorButtons: false,
      showReviveEliminatorButtons: false
    };
  }
}

/**
 * Creates and displays the HTML modal dialog for member management.
 */
function launchMemberPanel() {
  // Create an HTML output object from a separate HTML file.
  // This is cleaner than embedding a huge string in your .gs file.
  const html = HtmlService.createHtmlOutputFromFile('memberPanel')
      .setWidth(550) // Set a comfortable width for the dialog
      .setHeight(200); // And a reasonable height
  
  // Display it as a modal dialog. The user must interact with it before returning to the sheet.
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Management');
}

/**
 * Creates and displays the HTML modal dialog for member renaming
 */
function showRenamePanel() {
  const html = HtmlService.createHtmlOutputFromFile('renamePanel')
      .setWidth(400)
      .setHeight(280);  
  SpreadsheetApp.getUi().showModalDialog(html, 'Rename a Member');
}

/**
 * Processes the submitted member list, compares it to the previously saved
 * state, and performs the necessary add, rename, and delete operations on the spreadsheet.
 *
 * @param {Object} newMemberData The new state of the member list from the client.
 */
function processMemberSubmission(clientData) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const config = JSON.parse(docProps.getProperty('configuration'));
    let addedNames = [];
    // Fetch the "before" state
    const serverData = fetchProperties('members') || { memberOrder: [], members: {} };

    // Process the "after" state from the client, finalizing IDs
    const finalData = { memberOrder: [], members: {} };
    let currentWeek;
    clientData.memberOrder.forEach(id => {
      const memberDetails = clientData.members[id];
      
      if (id.startsWith('new_')) {
        const permanentId = generateUniqueId();
        finalData.memberOrder.push(permanentId);
        currentWeek = currentWeek || fetchWeek() || 1;
        finalData.members[permanentId] = createNewMember(
          memberDetails.name,
          memberDetails.paid,
          config,
          currentWeek
        );
      } else {
        // This is an existing member. Keep their permanent ID
        finalData.memberOrder.push(id);
        
        // Merge the old data with any new changes (like the 'paid' status)
        const existingData = serverData.members[id] || {};
        finalData.members[id] = {
          ...existingData, // Keep all old data (lives, revives, etc.)
          name: memberDetails.name, // In case of renames in the future
          paid: memberDetails.paid  // Update the paid status
        };
      }
    });
    
    // Perform Deletion Logic (This part is simplified)
    // Find any IDs that were in the original serverData but are NOT in the new finalData
    const initialIds = serverData.memberOrder || [];
    const finalIds = finalData.memberOrder;
    const deletedIds = initialIds.filter(id => !finalIds.includes(id));

    if (deletedIds.length > 0) {
        deletedIds.forEach(id => {
            const memberName = serverData.members[id]?.name || id;
            removeMemberFromSheet(memberName);
        });
    }

    // Identify Additions (for future use): Names in the new list not in the old one.
    if (addedNames.length > 0) {
      Logger.log(`Adding new members: ${addedNames}`);
      // Could call an `addMemberToSheet(name)` function here
    }
    
    // memberAddForm(addedNames);
    
    // Save the new, final, authoritative state
    saveProperties('members', finalData);
    
    return { success: true, message: 'Members updated successfully!' };

  } catch (err) {
    Logger.log(`⚠️ Error processing member submission: ${err.stack}`);
    throw new Error(`⚠️ Failed to update members. ${err.toString()}`);
  }
}

/**
 * Creates a complete, correctly structured object for a new member.
 *
 * @param {string} name The new member's name.
 * @param {boolean} isPaid The initial paid status.
 * @param {Object} config The main configuration object.
 * @param {number} joinWeek The week number the member is joining in.
 * @returns {Object} The complete new member object.
 */
function createNewMember(name, isPaid, config, joinWeek, extras) {
  // Create an array with (joinWeek - 1) empty slots for weeks they missed.
  const pastWeekPadding = Array(joinWeek > 1 ? joinWeek - 1 : 0).fill(null);

  const newMember = {
    name: name,
    paid: isPaid,
    active: true,
    joinDate: new Date().toISOString(),

    // Sign-up details, when the member came in through the sign-up form.
    // "name" is the team name (what shows on the WK sheets and in the weekly form);
    // "manager" is the actual person, which the payout tab needs.
    manager: (extras && extras.manager) || '',
    email: (extras && extras.email) || '',
    
    // Survivor Properties
    sR: [...pastWeekPadding],
    sP: [...pastWeekPadding],
    sE: [...pastWeekPadding],
    sL: [...pastWeekPadding], // The 'lives array' starts with padding for past weeks
    sO: null,
    
    // Eliminator Properties
    eR: [...pastWeekPadding],
    eP: [...pastWeekPadding],
    eE: [...pastWeekPadding],
    eL: [...pastWeekPadding],
    eO: null
  };

  // Append starting lives (or 1 as fallback) to the array IF the start week is the same as the join week, otherwise add 0
  if (joinWeek == (config.survivorStartWeek || 1)) {
    newMember.sL.push(parseInt(config.survivorLives, 10) || 1);
  } else {
    newMember.sL.push(0);
  }
  if (joinWeek == (config.eliminatorStartWeek || 1)) {
    newMember.eL.push(parseInt(config.eliminatorLives, 10) || 1);
  } else {
    newMember.eL.push(0);
  }
  
  return newMember;
}
 
 /**
 * Processes a rename submission. Finds the member by their
 * old name to get their unique ID, then updates the name property for that ID.
 *
 * @param {Object} data An object from the client with 'oldName' and 'newName' properties.
 */
function processRenameSubmission(data) {
  const oldName = data.oldName;
  const newName = data.newName.trim(); // Sanitize the new name

  // --- 1. Server-side validation ---
  if (!oldName || !newName || oldName.toLowerCase() === newName.toLowerCase()) {
    throw new Error("Invalid input. Please select a member and provide a different new name.");
  }
  
  const memberData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('members'));
  if (!memberData.members) {
    throw new Error("Could not find any member data to update.");
  }
  
  // a) Find the member's unique ID by their old name. This is a case-insensitive search.
  let memberIdToUpdate = null;
  let currentNames = []; // To check for duplicates
  
  for (const id in memberData.members) {
    const member = memberData.members[id];
    if (member.name.toLowerCase() === oldName.toLowerCase()) {
      memberIdToUpdate = id;
    }
    currentNames.push(member.name.toLowerCase());
  }

  // b) More validation now that we have the data
  if (!memberIdToUpdate) {
    throw new Error(`The member "${oldName}" could not be found. They may have already been renamed or deleted.`);
  }
  if (currentNames.includes(newName.toLowerCase())) {
    throw new Error(`The name "${newName}" already exists in the member list.`);
  }

  // --- 2. Perform the Update (Now incredibly simple) ---
  // We only need to change the name property of the specific member object.
  // The 'memberOrder' array of IDs does NOT need to be changed at all.
  memberData.members[memberIdToUpdate].name = newName;
  
  // --- 3. Save the updated object back to properties ---
  saveProperties('members', memberData);

  // --- 4. Run the function to update the name on all user-facing sheets ---
  // This function still works perfectly. It finds all instances of the old name
  // and replaces them with the new one.
  Logger.log(`✏️ Renaming sheet member names...`);
  renameMemberInSheet(oldName, newName);
  Logger.log(`✏️ Renaming database sheet member names...`);
  renameMemberInDatabaseSheet(oldName, newName);

  return { success: true, message: `Successfully renamed "${oldName}" to "${newName}".` };
}

/**
 * Finds all exact, case-sensitive matches of a member's name across all sheets
 * and removes the entire row where the name is found.
 *
 * @param {string} memberName The name of the member to remove.
 */
function removeMemberFromSheet(memberName) {
  if (!memberName) return; // Safety check

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  Logger.log(`Searching for rows to delete for member: "${memberName}"`);

  sheets.forEach(sheet => {
    const textFinder = sheet.createTextFinder(memberName)
      .matchEntireCell(true) // Crucial: Ensures "Ben" doesn't match "Benjamin"
      .matchCase(true);      // Ensures "Ben" doesn't match "ben"

    const foundCells = textFinder.findAll();
    
    if (foundCells.length > 0) {
      // We must delete from the bottom up to avoid shifting row indexes.
      foundCells.reverse().forEach(cell => {
        const row = cell.getRow();
        Logger.log(`Deleting row ${row} from sheet "${sheet.getName()}" because it contained "${memberName}"`);
        sheet.deleteRow(row);
      });
    }
  });
}

/**
 * Creates and displays the HTML modal dialog for member management.
 */
function launchSurvElimPanel() {
  // Create an HTML output object from a separate HTML file.
  // This is cleaner than embedding a huge string in your .gs file.
  const html = HtmlService.createHtmlOutputFromFile('survElimPanel')
      .setWidth(600)
      .setHeight(700);
  
  // Display it as a modal dialog. The user must interact with it before returning to the sheet.
  SpreadsheetApp.getUi().showModalDialog(html, 'Contest Manager');
}

/**
 * Retrieves all data needed for the survivor and eliminator manager panel.
 *
 * @returns {Object}
 */
function getSurvElimManagerData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();

    // 1. Fetch the member data as before.
    const members = JSON.parse(docProps.getProperty('members')) || { order: [], details: {} }; // Ensure a default object

    // 2. Fetch the main configuration
    const config = JSON.parse(docProps.getProperty('configuration')) || {}; // Ensure a default object
    Logger.log("Output from the getSurvelimManagerData: " + JSON.stringify({
      week: fetchWeek() || 1,
      memberData: members,
      config: config,
      leagueData: LEAGUE_DATA
    }))
    // 4. Return a single, bundled object with all the data the client needs.
    return {
      week: fetchWeek() || 1,
      memberData: members,
      config: config,
      leagueData: LEAGUE_DATA
    };

  } catch (error) {
    Logger.log('Error preparing survivor/eliminator data:', error);
    // Return a safe, default structure in case of any error.
    return {
      week: fetchWeek() || 1,
      memberData: { membersOrder: [], members: {} },
      config: config,
      leagueData: LEAGUE_DATA
    };
  }
}

/**
 * Saves the modified member data from the UI panel to Properties and updates the Sheets.
 * @param {Object} updatedMemberData The full memberData object from the HTML panel.
 */
function saveSurvElimData(updatedMemberData) {
  try {
    // 1. Save the JSON object to Document Properties
    saveProperties('members', updatedMemberData);

    // 2. Update the physical sheets (Survivor/Eliminator) so they match the new data
    // We pass null for 'week' to indicate a full refresh of all weeks, 
    // or you can just sync the current view.
    syncSurvElimDataToSheet(updatedMemberData);

    Logger.log(`✅ Member records updated/modified successfully`)
    return {
      success: true,
      message: `✅ Member records updated/modified successfully`
    };
  } catch (err) {
    Logger.log(`⚠️ Error in updating member suvivor/eliminator data: ${err.stack}`);
    throw new Error("⚠️ Failed to save data:\n" + err.message);
  }
}

/**
 * Refreshes the Spreadsheet picks and lives based on the current JSON data.
 * Used after manual edits in the Manager Panel.
 */
function syncSurvElimDataToSheet(memberData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const types = ['SURVIVOR', 'ELIMINATOR'];

  types.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const prefix = sheetName.charAt(0).toLowerCase(); // 's' or 'e'
    const nameRange = ss.getRangeByName(`${sheetName}_NAMES`);
    const namesOnSheet = nameRange.getValues().flat();
    
    // Map names to their IDs
    const nameToId = {};
    for (const id in memberData.members) {
      nameToId[memberData.members[id].name.toLowerCase().trim()] = id;
    }

    namesOnSheet.forEach((name, index) => {
      const memberId = nameToId[name.toString().toLowerCase().trim()];
      if (memberId) {
        const picks = memberData.members[memberId][prefix + 'P'] || [];
        if (picks.length > 0) {
          const rowIndex = index + nameRange.getRow();
          // Write the whole row of picks starting at Column 5 (Week 1)
          sheet.getRange(rowIndex, 5, 1, picks.length).setValues([picks]);
        }
      }
    });
  });
}

/**
 * Called by the HTML Panel to handle a revival.
 * Updates the data object and then refreshes the spreadsheet visuals.
 */
function recalculateAndReviveFromWeek(payload) {

  Logger.log(JSON.stringify(payload));
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Run the data logic (Update JSON)
    // processReviveMember handles the JSON and saveProperties
    const result = processReviveMember(payload);
    
    // 2. Refresh the Spreadsheet Visuals
    // We pass the updatedMemberData directly so updateSurvElimSheet doesn't have to re-fetch it
    const config = fetchProperties('configuration');
    updateSurvElimSheet(ss, config, result.updatedMemberData, payload.gameType);

    return result; // Return success to HTML to hide loader and alert user
  } catch (err) {
    Logger.log(`Error in recalculateAndReviveFromWeek: ${err.stack}`);
    throw new Error(err.message);
  }
}

/**
 * Fetches a simple, ordered list of all current member names.
 * This is used to populate the dropdown in the "Rename a Member" panel.
 * (This is the function you are likely referring to as fetchMembers).
 *
 * @returns {string[]} An array of current member names in the correct order.
 */
function fetchMembers() {
  try {
    // 1. Fetch the complete, authoritative members object.
    const memberData = fetchProperties('members');

    // 2. Safely get the order of member IDs. Default to an empty array if it doesn't exist.
    const memberOrder = memberData.memberOrder || [];
    const members = memberData.members || {};

    // 3. Map the array of IDs to an array of names.
    const names = memberOrder.map(id => {
      // For each ID, look up the corresponding member's name.
      // The optional chaining (?.) prevents errors if an ID has no matching member object.
      return members[id]?.name;
    }).filter(name => name); // 4. Filter out any null or undefined names.

    return names;

  } catch (err) {
    Logger.log(`⚠️ Error in getMemberNames: ${err.stack}`);
    return []; // Always return an array, even on failure.
  }
}

/**
 * Finds all exact, case-sensitive matches of an old member name and replaces
 * it with the new name.
 *
 * @param {string} oldName The original name to find.
 * @param {string} newName The new name to replace it with.
 */
function renameMemberInSheet(oldName, newName) {
  if (!oldName || !newName || oldName === newName) return; // Safety check

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log(`Searching for cells to rename "${oldName}" to "${newName}"`);

  // TextFinder can operate on the entire spreadsheet at once.
  const textFinder = ss.createTextFinder(oldName)
    .matchEntireCell(true) // Crucial: Fulfills your requirement
    .matchCase(true);

  // replaceAllWith is a single, efficient operation.
  const cellsReplaced = textFinder.replaceAllWith(newName);
  Logger.log(`Replaced ${cellsReplaced} instances of "${oldName}".`);
}

/**
 * Finds and replaces a member's name across all relevant columns
 * in the private backend response spreadsheet.
 *
 * @param {string} oldName The original name to find.
 * @param {string} newName The new name to replace it with.
 */
function renameMemberInDatabaseSheet(oldName, newName) {
  try {
    const dbSheet = getDatabaseSheet(); // Your existing helper to get the backend Spreadsheet
    if (!dbSheet) {
      Logger.log(`❌ Database sheet not found, skipping rename operation there.`);
      return;
    }

    // Find the columns that contain member names. We'll use our regex helper for this.
    const allSheets = dbSheet.getSheets();
    
    allSheets.forEach(sheet => {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const { nameCol, newNameCol } = findNameColumns(headers); // Your existing regex helper

      // Replace in the "Select Your Name" column
      if (nameCol > -1) {
        const range = sheet.getRange(2, nameCol + 1, sheet.getLastRow());
        range.createTextFinder(oldName).matchEntireCell(true).replaceAllWith(newName);
      }
      
      // Replace in the "Enter Your Name" column
      if (newNameCol > -1) {
        const range = sheet.getRange(2, newNameCol + 1, sheet.getLastRow());
        range.createTextFinder(oldName).matchEntireCell(true).replaceAllWith(newName);
      }
      
      // Replace in the unique "Survivor Winner Pick (Name)" columns
      headers.forEach((header, index) => {
        if (header.includes(`(${oldName})`)) {
          const newHeader = header.replace(`(${oldName})`, `(${newName})`);
          sheet.getRange(1, index + 1).setValue(newHeader);
        }
      });
    });
    
    Logger.log(`✅ Successfully performed rename operations for "${oldName}" in the backend database sheet.`);

  } catch (err) {
    // Primary rename (in properties) succeeded
    // This is a secondary cleanup task
    Logger.log(`⚠️ Could not complete rename in the database sheet for "${oldName}". ${err.stack}`);
  }
}

/**
 * Looks for form for the week, checks what members are listed within the "eligible" entrants and adds new ones
 * 
 * @param {array} names of new players to add
 * @param {integer} week number to review
 */
function memberAddForm(names,week){
  const config = fetchConfiguration();
  const ss = fetchSpreadsheet();

  if (week == null) {
    week = fetchWeek();
  }
  if (typeof names == 'string') {
    names = [names];
  } else if (names == null) {
    names = ['New User'];
  }
  let nameQuestion, gotoPage, newUserPage, found = false;
  try {
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValue(); // FLAG
    if (formId) {
      let form = FormApp.openById(formId);
      const items = form.getItems();
      for (let a = 0; a < items.length; a++) {
        if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
          nameQuestion = items[a];
          found = true;
        } else if (items[a].getType() == 'PAGE_BREAK'){
          let pageBreakItem = items[a].asPageBreakItem();
          let pageTitle = pageBreakItem.getTitle();
          if (pageTitle == 'Survivor Start') {
            gotoPage = pageBreakItem;
          } else if (pageTitle == 'New User') {
            newUserPage = pageBreakItem;
          }
        }
      }
      if (found && nameQuestion) {
        let newChoice, choices = nameQuestion.asListItem().getChoices();
        if (config.survivorInclude && survivorStart == week) { // FLAG
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                Logger.log(`New user "${names[a]}" is redirected to the "${newUserPage.getTitle()}" Form page`);
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],gotoPage);
                Logger.log(`New user "${names[a]}" is redirected to the "${gotoPage.getTitle()}" Form page`);
              }
              choices.unshift(newChoice);
              
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast(`Issue locating survivor start question, you may need to add member manually`,`❗ ISSUE`);
            Logger.log(`⚠️ Error with "memberAdd": ${err.stack}`);
          }
        } else {
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                choices.unshift(newChoice);
                Logger.log(`↪️ New user "${names[a]}" is redirected to the "${newUserPage.getTitle()}" Form page`);
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],FormApp.PageNavigationType.SUBMIT);
                choices.unshift(newChoice);
                Logger.log(`↪️ New user "${names[a]}" is redirected to the submit Form page`);
              }
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast('Issue locating submit form value, you may need to add member manually');
            Logger.log(`⚠️ Error with "memberAdd": ${err.stack}`);
          }
        }
      }
    } else {
      Logger.log(`⏭ No form created yet for week ${week}, skipping addition of ${names} to form.`);
    }
  }
  catch (err) {
    Logger.log(`⚠️ Error with "memberAdd": ${err.stack}`);
    ss.toast(`Unable to add ${names} to the form.`,`⚠️ ERROR ADDING MEMBER(S) TO FORM`);
  }
}


// ============================================================================================================================================
// API PULLING
// ============================================================================================================================================

// SEASON INFORMATION FUNCTIONS
//------------------------------------------------------------------------
/** 
 * Year fetching from configuration if available, if none available, fetches from ESPN API
 * 
 * @returns {int} the four digit year value of the season
 */
function fetchYear(apiPull) {
  const docProps = PropertiesService.getDocumentProperties();
  let config = JSON.parse(docProps.getProperty('configuration')) || {};
  const yearRegEx = new RegExp(/[0-9]{4}/,'g');
  let yearInvalid = true;
  if (config && !apiPull) {
    year = config.year;
    if (year) {
      return (parseInt(year).toFixed(0));;
    }
  }
  let success = false;
  if (apiPull) {
    Logger.log(`📡 API Pull requested to fetch year`);
  } else {
    Logger.log(`❌ No year currently recorded for league, fetching from ESPN API...`)
  }
  try {
    year = espnFetchJson(SCOREBOARD).season.year.toString();
    if (year) {
      yearInvalid = !yearRegEx.test(year);
      if (yearInvalid) {
        Logger.log(`⚠️ API WARNING: Year value of "${year}" pulled from API but was not a valid year, moving on to manual submission`);
      } else {
        Logger.log(`✅ API SUCCESS: Pulled year value of "${year}" from season information from ESPN API`)
        success = true;
      }        
    }
  } catch (err) {
    Logger.log(`API FAILURE: Unable to pull API data, moving to prompt for user... (${err})`);
  }
  if (!success) { 
    const ui = fetchUi();
    let retry = true;
    let yearPrompt = ui.prompt(`Year Entry`, `Please submit the year in a YYYY format to set the league's year for this season:`, ui.ButtonSet.OK_CANCEL);
    while (retry && yearInvalid) {
      Logger.log(JSON.stringify(yearPrompt));
      Logger.log(yearPrompt.getResponseText())
      if (yearPrompt.getSelectedButton() == ui.Button.OK) {
        year = yearPrompt.getResponseText();
        Logger.log(`Received year entry of "${year}"`);
        yearInvalid = !yearRegEx.test(year);
        if (yearInvalid) {
          yearPrompt = ui.prompt(`Retry Year Entry`, `That wasn't a valid submission, please submit the year in a YYYY format to set the league's year for this season:`, ui.ButtonSet.OK_CANCEL);
        }
      } else {
        showToast(`User canceled manual entry of year`);
        retry = false;
      }
    }
  }
  if (!yearInvalid) {
    Logger.log(`✏️ Storing year value in "configuration" property...`)
    config.year = year;
    try {
      saveProperties('configuration',config);
      Logger.log(`✅ SUCCESS: Stored year key of "${year}" within the document properties.`)
      return (parseInt(year).toFixed(0));
    } catch (err) {
      Logger.log(`⚠️ ERROR: Issue storing year key: ${err.stack}`);
      return null;
    }
  } else {
    return null;
  }
}

// FETCH CURRENT WEEK
function fetchWeek(negative,current) {
  let weeks, week, advance = 0;
  try {
    const obj = espnFetchJson(SCOREBOARD);
    let season = obj.season.type;
    obj.leagues[0].calendar.forEach(entry => {
      if (entry.value == season) {
        weeks = entry.entries.length;
      }
    });
    obj.events.forEach(event => {
      if (event.status.type.state != 'pre' && !current) {
        advance = 1; // At least one game has started and therefore the script will prompt for the next week
      }
    });
    let name;
    switch (season) {
      case 1:
        name = 'Preseason';
        week = obj.week.number - (weeks + 1);
        break;
      case 2: 
        name = 'Regular Season';
        week = obj.week.number + advance;
        break;
      case 3:
        name = 'Postseason';
        week = obj.week.number + obj.leagues[0].calendar[1].entries.length + advance;
        break;
    }
    Logger.log(name + ' is currently active with ' + weeks + ' weeks in total, current week is: ' + week); 
    if (negative) {
      
      return week;
    } else {
      week = week <= 0 ? 1 : week;
      return week;
    }
  }
  catch (err) {
    Logger.log(`⚠️ ESPN API has an issue right now ${err.stack}`);
    return null;
  }
}

// SCOREBOARD WEEK ENDPOINT
function fetchScoreboardEndpoint(week, leg) {
  let current;
  
  // Assume regular season if only week is provided
  if (week && !leg) {
    leg = 2;
  }

  // Fetch current data only if week or leg is still missing
  if (!week || !leg) {
    current = espnFetchJson(SCOREBOARD);
  }
  
  week = week || current.week.number;
  leg = leg || current.season.type;
  
  // Clamp leg to valid range (1-3)
  leg = leg > 3 ? 3 : (leg < 1 ? 1 : leg);
  
  // Convert to postseason format if week exceeds regular season
  if (week > REGULAR_SEASON) {
    week = week - REGULAR_SEASON;
    // Only change leg to postseason if it's currently regular season
    if (leg === 2) {
      Logger.log(`⏭ Input provided was beyond the regular season, directing API to Post Season.`);
      leg = 3;
    }
  }

  if (week && leg) {
    return `${SCOREBOARD}?seasontype=${leg}&week=${week}`;
  } else {
    Logger.log(`⚠️ Issue with inputs or fetching current week/leg via "fetchScoreboardEndpoint" function, returning base scoreboard.`);
    return SCOREBOARD;
  }
}

// ESPN FUNCTIONS
//------------------------------------------------------------------------
// ESPN TEAMS - Fetches the ESPN-available API data on NFL teams
function fetchTeamsESPN(year) {
  if (year == undefined) {
    year = fetchYear();
  }

  let obj = {};
  try {
    let string = schedulePrefix + year + scheduleSuffix;
    Logger.log(`🔎 Fetching JSON content from: ${string}`);
    obj = espnFetchJson(string);
    let objTeams = obj.settings.proTeams;
    return objTeams;
  }
  catch (err) {
    Logger.log(`⚠️ ESPN API has an issue right now: ${err.stack}`);
  }  
}

// NFL TEAM INFO - script to fetch all NFL data for teams - auto for setting up trigger allows for boolean entry in column near the end
function fetchSchedule(ss,year,currentWeek,auto,overwrite) {
  // Calls the linked spreadsheet
  const timeFetched = new Date();
  ss = fetchSpreadsheet(ss);
  let all = false;
  if (currentWeek == undefined || currentWeek == null) {
    currentWeek = fetchWeek(null,true);
    all = true;
    ss.toast(`Fetching complete schedule data for the ${LEAGUE}`,`📅 FETCHING SCHEDULE`);
  } else {
    ss.toast(`Fetching only data for week ${currentWeek}, if available.`,`📅 FETCHING WEEK ${currentWeek}`);
  }
  // Declaration of script variables
  if (year == undefined || year == null) {
    year = fetchYear();
  }
  const objTeams = fetchTeamsESPN(year);
  if (!objTeams) {
    Logger.log("❌ fetchTeamsESPN returned undefined. Aborting schedule fetch.");
    ss.toast("Could not retrieve team data from ESPN.", "❌ API ERROR");
    return;
  }
  const teamsLen = objTeams.length;
  let headers = ['week','date','day','hour','minute','dayName','awayTeam','homeTeam','awayTeamLocation','awayTeamName','homeTeamLocation','homeTeamName','type','divisional','division','overUnder','spread','spreadAutoFetched','timeFetched'];
  const sheetName = LEAGUE;
  const overUnderIdx = headers.indexOf('overUnder');
  const spreadIdx = headers.indexOf('spread');
  const awayTeamIdx = headers.indexOf('awayTeam');
  const homeTeamIdx = headers.indexOf('homeTeam');
  const spreadAutoIdx = headers.indexOf('spreadAutoFetched');
  const timeFetchedIdx = headers.indexOf('timeFetched');

  let sheet, range, abbr, name, arr = [], nfl = [],espnId = [], espnAbbr = [], espnName = [], espnLocation = [], location = [], ids = [], abbrs = []; 

  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    if(objTeams[a].id != 0 ) {
      abbr = objTeams[a].abbrev.toUpperCase();
      name = objTeams[a].name;
      location = objTeams[a].location;
      espnId.push(objTeams[a].id);
      espnAbbr.push(abbr);
      espnName.push(name);
      espnLocation.push(location);
    }
  }

  for (let a = 0 ; a < espnId.length ; a++ ) {
    ids.push(espnId[a].toFixed(0));
    abbrs.push(espnAbbr[a]);
  }

  // Declaration of variables
  let schedule = [], home = [], dates = [], allDates = [], hours = [], allHours = [], minutes = [], allMinutes = [], byeIndex, id, date, hour, minute, weeks = Object.keys(objTeams[0].proGamesByScoringPeriod).length;
  if ( objTeams[0].byeWeek > 0 ) {
    weeks++;
  }

  location = [];
  
  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    home = [];
    dates = [];
    hours = [];
    minutes = [];
    byeIndex = objTeams[a].byeWeek.toFixed(0);
    if ( byeIndex != 0 ) {
      id = objTeams[a].id.toFixed(0);
      arr.push(abbrs[ids.indexOf(id)]);
      home.push(abbrs[ids.indexOf(id)]);
      dates.push(abbrs[ids.indexOf(id)]);
      hours.push(abbrs[ids.indexOf(id)]);
      minutes.push(abbrs[ids.indexOf(id)]);
      for (let b = 1 ; b <= weeks ; b++ ) {
        if ( b == byeIndex ) {
          arr.push('BYE');
          home.push('BYE');
          dates.push('BYE');
          hours.push('BYE');
          minutes.push('BYE');
        } else {
          if ( objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0) === id ) {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].awayProTeamId.toFixed(0))]);
            home.push(1);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          } else {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0))]);
            home.push(0);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          }
        }
      }
      schedule.push(arr);
      location.push(home);
      allDates.push(dates);
      allHours.push(hours);
      allMinutes.push(minutes);
    }
  }
  
  // This section creates a nice table to be used for lookups and queries about NFL season
  let week, awayTeam, awayTeamName, awayTeamLocation, homeTeam, homeTeamName, homeTeamLocation, day, dayName, divisional, division, scheduleData = [];
  
  // Create an array of matchups per week where index of 0 is equivalent to week 1 and so forth
  let matchupsPerWeek = Array(WEEKS).fill(0);
  arr = [];
  let weekArr = [];
  for (let b = 0; b < (teamsLen - 1); b++ ) {
    for ( let c = 1; c <= weeks; c++ ) {
      if (location[b][c] == 1) {
        week = c;
        awayTeam = schedule[b][c];
        awayTeamName = espnName[espnAbbr.indexOf(awayTeam)];
        awayTeamLocation = espnLocation[espnAbbr.indexOf(awayTeam)];
        homeTeam = schedule[b][0];
        homeTeamName = espnName[espnAbbr.indexOf(homeTeam)];
        homeTeamLocation = espnLocation[espnAbbr.indexOf(homeTeam)];
        date = allDates[b][c];
        hour = allHours[b][c];
        minute = allMinutes[b][c];
        day = date.getDay();
        // Uses globalVariables.gs variable to determine day name and assign offset index
        dayName = DAY[day].name;
        day = DAY[day].index;
        divisional = LEAGUE_DATA[homeTeam].division_opponents.indexOf(awayTeam) > -1 ? 1 : 0;
        division = divisional == 1 ? LEAGUE_DATA[homeTeam].division : '';

        arr = [
          week,
          date,
          day,
          hour,
          minute,
          dayName,
          awayTeam,
          homeTeam,
          awayTeamLocation,
          awayTeamName,
          homeTeamLocation,
          homeTeamName,
          WEEKNAME.hasOwnProperty(c) ? WEEKNAME[c].name : 'Regular Season', // type
          divisional,
          division,
          '', // Placeholder for overUnder
          '', // Placeholder for spread
          '', // Placeholder for spreadAutoFetched
          timeFetched
        ];
        matchupsPerWeek[week-1] = matchupsPerWeek[week-1] + 1;
        scheduleData.push(arr);
      }
    }
  }

  scheduleData = scheduleData.sort((a,b) => a[1] - b[1]);
    
  for (let a = 0; a < scheduleData.length; a++) {
    weekArr.push(scheduleData[a][0]);
  }
  // Add the playoff schedule to that array of matchups per week
  Object.keys(WEEKNAME).forEach(weekNum => {
    matchupsPerWeek[weekNum-1] = WEEKNAME[weekNum].matchups;
    for (let a = 0; a < matchupsPerWeek[weekNum-1]; a++) {
      weekArr.push(parseInt(weekNum));
    }
  });

  // Create indexing array of when weeks begin and end
  let rowIndex = 2;
  let startingRow = Array(WEEKS).fill(0);
  for (let a = 1; a < startingRow.length; a++) {
    let start = 0;
    for (let b = 0; b < a; b++) {
      start = start + matchupsPerWeek[b];
    }
    startingRow[a] = 2 + start;
  }


  // Sheet formatting & Range Setting =========================
  sheet = ss.getActiveSheet();
  if ( sheet.getSheetName() == 'Sheet1' && ss.getSheetByName(sheetName) == null) {
    sheet.clear();
    sheet.setName(sheetName);
  }
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(scheduleTabColor);

  adjustColumns(sheet,headers.length);

  sheet.setColumnWidths(1,headers.length,30);
  sheet.setColumnWidth(headers.indexOf('date')+1,60);
  sheet.setColumnWidth(headers.indexOf('dayName')+1,60);
  sheet.setColumnWidths(headers.indexOf('awayTeamLocation')+1,4,80); // All Locations & Team Names
  sheet.setColumnWidth(headers.indexOf('type')+1,110);
  sheet.setColumnWidth(headers.indexOf('division')+1,60);
  sheet.setColumnWidth(headers.indexOf('spread')+1,60);
  sheet.setColumnWidth(headers.indexOf('timeFetched')+1,110);
  range = sheet.getRange(1,1,1,headers.length);
  range.setValues([headers]);
  ss.setNamedRange(sheetName+'_HEADERS',range);

  range = sheet.getRange(1,1,weekArr.length+1,headers.length);
  range.setFontSize(8);
  range.setVerticalAlignment('middle');  
 
  ss.setNamedRange(sheetName,range);
  let rangeData = sheet.getRange(2,1,weekArr.length,headers.length);

  rangeData.setHorizontalAlignment('left');
  sheet.getRange(1,3).setNote('-4: Wednesday, -3: Thursday, -2: Friday, -1: Saturday, 0: Sunday, 1: Monday, 2: Tuesday');
  
  // Fetch existing sheet values first so we have them loaded
  let existingData = rangeData.getValues();
  const regexOverUnder = new RegExp(/^[0-9\.]+$/);
  const regexSpread = new RegExp(/^[A-Z]{2,3}\ \-[0-9\.]+$/);
  let existing = {};
  
  for (let a = 0; a < existingData.length; a++) {
    const row = existingData[a];
    const weekNum = row[0];
    const overUnderVal = row[overUnderIdx];
    const spreadVal = row[spreadIdx];

    if (regexOverUnder.test(overUnderVal) || regexSpread.test(spreadVal) || weekNum > REGULAR_SEASON) {
      const matchup = `${row[awayTeamIdx]}@${row[homeTeamIdx]}`;
      existing[weekNum] = existing[weekNum] || {};
      existing[weekNum][matchup] = {
        row,
        placed: false,
        auto: row[spreadAutoIdx],
        timeFetched: row[timeFetchedIdx],
        ...(overUnderVal && { overUnder: overUnderVal }),
        ...(spreadVal && { spread: spreadVal })
      };
    }
  }

  // Restore the existing spreadsheet values into scheduleData right away
  for (let a = 0; a < scheduleData.length; a++) {
    let weekNum = scheduleData[a][0];
    let matchup = `${scheduleData[a][awayTeamIdx]}@${scheduleData[a][homeTeamIdx]}`;
    if (existing[weekNum] && existing[weekNum].hasOwnProperty(matchup)) {
      let ext = existing[weekNum][matchup];
      if (ext.overUnder) scheduleData[a][overUnderIdx] = ext.overUnder;
      if (ext.spread) scheduleData[a][spreadIdx] = ext.spread;
      scheduleData[a][spreadAutoIdx] = ext.auto;
      scheduleData[a][timeFetchedIdx] = ext.timeFetched;
    }
  }

  // Set named ranges for weekly home and away teams to compare for survivor status
  awayTeam = headers.indexOf('awayTeam')+1;
  homeTeam = headers.indexOf('homeTeam')+1;
  ss.setNamedRange(`${LEAGUE}_MATCHUPS_HEADERS`,sheet.getRange(1,1,1,headers.length));
  for (let a = 0; a < WEEKS; a++) {
    if (matchupsPerWeek[a] > 0) {
      try {
        let start = weekArr.indexOf(a+1)+2;
        let len = matchupsPerWeek[a];
        ss.setNamedRange(`${LEAGUE}_AWAY_${a+1}`,sheet.getRange(start,awayTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_HOME_${a+1}`,sheet.getRange(start,homeTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_MATCHUPS_${a+1}`,sheet.getRange(start,1,len,headers.length));
      }
      catch (err) {
        Logger.log(`No data entered or available for week ${a} in the spreadsheet`);
        Logger.log(err.stack);
      }
    } else {
      Logger.log(`No matchups in week ${a}`);
    }
  }
  // Sheet formatting =========================

  // Set of loops to create blank entries for playoff schedule
  const blankRow = new Array(headers.length).fill('');
  Object.keys(WEEKNAME).forEach(weekNum => {
    const weekInt = parseInt(weekNum);
    if (weekInt > REGULAR_SEASON) {
      for (let b = 0; b < WEEKNAME[weekNum].matchups; b++) {
        let newRow = [...blankRow];
        newRow[0] = weekInt; // Replace first value with week number
        scheduleData.push(newRow);
      }
    }
  });

  // --- Scoreboard Retrieval Block ---

  // 1. Determine target week: If we are in the preseason (< 1), default to grabbing Week 1 spreads
  let targetWeek = currentWeek < 1 ? 1 : currentWeek;
  
  // Construct the Scoreboard URL dynamically based on the target week / postseason
  let scoreboardUrl = `${SCOREBOARD}?week=${currentWeek > REGULAR_SEASON ? currentWeek - REGULAR_SEASON : targetWeek}&seasontype=${currentWeek > REGULAR_SEASON ? 3 : 2}`;
  Logger.log(`Fetching scoreboard from ${scoreboardUrl}`);

  let scoreboardData = [];
  try {
    const response = espnFetch(scoreboardUrl);
    const obj = JSON.parse(response.getContentText());
    
    for (let event = 0; event < obj.events.length; event++) {
      date = new Date(obj.events[event].date);
      hour = date.getHours();
      minute = date.getMinutes();
      day = date.getDay();
      
      const away = obj.events[event].competitions[0].competitors.find(x => x.homeAway === 'away').team;
      const home = obj.events[event].competitions[0].competitors.find(x => x.homeAway === 'home').team;
      
      divisional = LEAGUE_DATA[home.abbreviation].division_opponents.indexOf(away.abbreviation) > -1 ? 1 : 0;
      division = divisional == 1 ? LEAGUE_DATA[home.abbreviation].division : '';
      
      // Safely extract the odds object if it exists
      let odds = obj.events[event].competitions[0].odds?.[0] || {};
      let overUnder = odds.overUnder || '';
      let details = odds.details || '';

      let arr = [
        targetWeek, // Map to targetWeek (1 or postseason target) instead of the negative preseason week
        date,
        DAY[day].index,
        hour,
        minute,
        DAY[day].name,
        away.abbreviation,
        home.abbreviation,
        away.location,
        away.name,
        home.location,
        home.name,
        WEEKNAME.hasOwnProperty(targetWeek) ? WEEKNAME[targetWeek].name : 'Regular Season',
        divisional,
        division,
        overUnder,
        details,
        auto ? 1 : 0,
        timeFetched
      ];
      scoreboardData.push(arr);
    }
  } catch (err) {
    Logger.log(`⚠️ Unable to fetch scoreboard details for Week ${targetWeek}: ${err.message}`);
  }

  // 2. Reconcile differences and merge at the matchup level
  let conflicts = [];
  // Retain a deep copy of raw, untouched scoreboard values
  let rawScoreboardData = scoreboardData.map(row => [...row]);

  if (scoreboardData.length > 0) {
    for (let s = 0; s < scoreboardData.length; s++) {
      let sbRow = scoreboardData[s];
      let matchupKey = `${sbRow[awayTeamIdx]}@${sbRow[homeTeamIdx]}`;
      
      // Match exact matchups for the target week
      let matchIdx = scheduleData.findIndex(row => 
        row[0] === targetWeek && 
        row[awayTeamIdx] === sbRow[awayTeamIdx] &&
        row[homeTeamIdx] === sbRow[homeTeamIdx]
      );
      
      if (matchIdx !== -1) {
        let existingRow = scheduleData[matchIdx];
        let oldOverUnder = existingRow[overUnderIdx];
        let oldSpread = existingRow[spreadIdx];
        let newOverUnder = sbRow[overUnderIdx];
        let newSpread = sbRow[spreadIdx];
        
        // Safety: If the API does not currently have betting lines, preserve the old values
        if (newOverUnder === '') {
          sbRow[overUnderIdx] = oldOverUnder;
          rawScoreboardData[s][overUnderIdx] = oldOverUnder;
          newOverUnder = oldOverUnder;
        }
        if (newSpread === '') {
          sbRow[spreadIdx] = oldSpread;
          rawScoreboardData[s][spreadIdx] = oldSpread;
          newSpread = oldSpread;
        }
        
        // Log discrepancies and queue them for display
        if (oldOverUnder !== newOverUnder || oldSpread !== newSpread) {
          Logger.log(`🔄 Conflict detected for ${matchupKey} (Week ${targetWeek}):`);
          Logger.log(`   Old Value -> O/U: ${oldOverUnder || 'None'}, Spread: ${oldSpread || 'None'}`);
          Logger.log(`   New Value -> O/U: ${newOverUnder || 'None'}, Spread: ${newSpread || 'None'}`);
          
          conflicts.push({
            matchup: matchupKey,
            oldOU: oldOverUnder || 'None',
            newOU: newOverUnder || 'None',
            oldSpread: oldSpread || 'None',
            newSpread: newSpread || 'None'
          });
        }
        
        // Safeguard: Protect manual overrides temporarily in scheduleData
        if (!overwrite && (oldOverUnder !== '' || oldSpread !== '')) {
          sbRow[overUnderIdx] = oldOverUnder;
          sbRow[spreadIdx] = oldSpread;
          sbRow[spreadAutoIdx] = existingRow[spreadAutoIdx];
        }
        
        scheduleData[matchIdx] = sbRow;
      }
    }
  }

  // 3. Prompt user dynamically for the targetWeek overrides
  if (!overwrite && conflicts.length > 0) {
    let ui = fetchUi();
    let message = `Found ${conflicts.length} matchups with conflicting betting lines for Week ${targetWeek}.\n\n`;
    
    let displayLimit = 5;
    for (let i = 0; i < Math.min(conflicts.length, displayLimit); i++) {
      let c = conflicts[i];
      message += `🏈 ${c.matchup}:\n`;
      message += `   Current -> O/U: ${c.oldOU}, Spread: ${c.oldSpread}\n`;
      message += `   Incoming -> O/U: ${c.newOU}, Spread: ${c.newSpread}\n\n`;
    }
    
    if (conflicts.length > displayLimit) {
      message += `...and ${conflicts.length - displayLimit} other matchup conflicts.\n\n`;
    }
    
    message += `Would you like to overwrite your existing data with these new API values?`;
    
    let replaceAlert = ui.alert(`⚠️ CONFLICTING BETTING LINES FOUND`, message, ui.ButtonSet.YES_NO_CANCEL);
    
    if (replaceAlert === ui.Button.YES) {
      // Overwrite: Apply the raw scoreboard details back into scheduleData
      for (let s = 0; s < rawScoreboardData.length; s++) {
        let rawRow = rawScoreboardData[s];
        let matchIdx = scheduleData.findIndex(row => 
          row[0] === targetWeek && 
          row[awayTeamIdx] === rawRow[awayTeamIdx] && 
          row[homeTeamIdx] === rawRow[homeTeamIdx]
        );
        if (matchIdx !== -1) {
          scheduleData[matchIdx][overUnderIdx] = rawRow[overUnderIdx];
          scheduleData[matchIdx][spreadIdx] = rawRow[spreadIdx];
          scheduleData[matchIdx][spreadAutoIdx] = auto ? 1 : 0;
          scheduleData[matchIdx][timeFetchedIdx] = timeFetched;
        }
      }
    }
  }

  // Checking for postseason empty slots within recently pulled data
  let missingMatchups = {};
  if (currentWeek > REGULAR_SEASON) {
    for (let a = 0; a < scheduleData.length; a++) {
      let scheduleDataWeek = scheduleData[a][0];
      if (scheduleDataWeek > REGULAR_SEASON) {
        if (scheduleData[a][headers.indexOf('awayTeam')] == '' || scheduleData[a][headers.indexOf('homeTeam')] == '') {
          missingMatchups[scheduleDataWeek] = missingMatchups[scheduleDataWeek] || {};
          missingMatchups[scheduleDataWeek].rows = missingMatchups[scheduleDataWeek].rows || [];
          missingMatchups[scheduleDataWeek].rows.push(a);
          missingMatchups[scheduleDataWeek].count = missingMatchups[scheduleDataWeek].count + 1 || 1;
        }
      }
    }
  }

  Object.keys(missingMatchups).forEach(week => {
    if (missingMatchups[week].count == matchupsPerWeek[week-1]) {
      try {
        Object.keys(existing[week]).forEach(matchup => {
          if (!existing[week][matchup].placed) {
            scheduleData[missingMatchups[week].rows[0]] = existing[week][matchup].row;
            existing[week][matchup].placed = true;
            missingMatchups[week].rows.splice(0,1);
          } else {
            Logger.log(`Already placed week ${week} matchup of ${matchup}.`);
          }
        });
      } catch (err) {
        Logger.log(`🔎 Unable to find any data week ${week} entry matchups, skipping...`);
      }
    } else {
      let emptyRows = [];
      let knownMatchups = [];
      for (let a = 0; a < scheduleData.length; a++) {
        if (scheduleData[a][0] == week) {
          if (scheduleData[a][headers.indexOf('awayTeam')] != '' && scheduleData[a][headers.indexOf('homeTeam')] != '') {
            knownMatchups.push(scheduleData[a]);
          } else {
            emptyRows.push(a);
          }
        }
      }
      for (let a = 0; a < knownMatchups.length; a++) {
        let weekNum = knownMatchups[a][0];
        let matchupKey = `${knownMatchups[a][headers.indexOf('awayTeam')]}@${knownMatchups[a][headers.indexOf('homeTeam')]}`;
        
        if (existing[weekNum] && existing[weekNum].hasOwnProperty(matchupKey)) {
          existing[weekNum][matchupKey].placed = true;
        }
      }
      Object.keys(existing[week]).forEach(matchup => {
        if (!existing[week][matchup].placed) {
          scheduleData.splice(emptyRows[0],1,existing[week][matchup].row);
          emptyRows.shift();
          existing[week][matchup].placed = true;
        }
      });
    }
  });

  for (let a = 0; a < scheduleData.length; a++ ) {
    try {
      let scheduleDataWeek = scheduleData[a][0];
      if (existing.hasOwnProperty(scheduleDataWeek)) {     
        if (existing[scheduleDataWeek].hasOwnProperty('row')) {
          Logger.log(`Replacing ${scheduleData[a]} with object data: ${existing[scheduleDataWeek].row}`);
          scheduleData.splice(a,1,existing[scheduleDataWeek].row);
        }
      }
    } catch (err) {
      // Logger.log(`No existing data for week ${a}`)
    }
  }

  let rows = scheduleData.length + 1;
  let columns = scheduleData[0].length;
  
  // utilities.gs functions to remove/add rows that are blank
  adjustRows(sheet,rows);
  adjustColumns(sheet,columns);

  try {
    rangeData.setValues(scheduleData);
    sheet.protect().setDescription(sheetName);
    try {
      sheet.hideSheet();
    }
    catch (err){
      Logger.log(`❗ fetchSchedule hiding: Couldn't hide sheet as no other sheets exist`);
    }
      ss.toast(`Imported all ${LEAGUE} schedule data`,`✅ SUCCESS`);
  } catch (err) {
    Logger.log(`❗ Issue placing the schedule data: ${err.message}`);
    ss.toast(`fetchSchedule failed to import ${LEAGUE} schedule data`, `❗ FAILURE`);
  }  
}

/**
 * Updates the spread and over/under data.
 * If all games in the current week are completed, it automatically fetches data for the next week.
 *
* @param {boolean} headless - If true, runs without user prompts.
 * @param {number} [targetWeek] - Optional. A specific week number to fetch data for, bypassing auto-detection.
 */
function fetchLatestSpreadsForWeek(headless, targetWeek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = headless ? null : SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(LEAGUE);

  if (!sheet) {
    if (ui) {
      ui.alert(`⚠️ Error`, `"${LEAGUE}" schedule sheet not found, bringing in schedule data now. Re-run this script to fetch spreads after ${LEAGUE} sheet creation has completed.`, ui.ButtonSet.OK);
      fetchSchedule(ss); // Assumes fetchSchedule is available to build the sheet
    } else {
      Logger.log(`❗ Attempted to fetch current spreads in headless mode and was unable to find the ${LEAGUE} schedule sheet.`);
      return;
    }
  }

  try {
    let weekToUpdate;
    let year;
    let dataToProcess;

    if (targetWeek) {
      // If a specific week is provided, use it directly
      weekToUpdate = targetWeek;
      if (!headless) ss.toast(`Checking for data in Week ${weekToUpdate}...`, `🎯 Targeting Week ${weekToUpdate}`);
      const response = espnFetch(`${SCOREBOARD}?week=${weekToUpdate}`);
      dataToProcess = JSON.parse(response.getContentText());
      year = dataToProcess.season.year;
    } else {
      // --- Auto-detection logic ---
      const initialResponse = espnFetch(SCOREBOARD);
      dataToProcess = JSON.parse(initialResponse.getContentText());
      
      let currentWeek = dataToProcess.week.number;
      weekToUpdate = currentWeek;
      year = dataToProcess.season.year;

      const allGamesCompleted = dataToProcess.events && dataToProcess.events.length > 0 && 
                                dataToProcess.events.every(event => event.status.type.completed === true);

      if (allGamesCompleted) {
        if (!headless) ss.toast(`Week ${currentWeek} is complete. Targeting next week...`, `⏩ Week ${currentWeek + 1}`);
        weekToUpdate = currentWeek + 1;
        const nextWeekResponse = espnFetch(`${SCOREBOARD}?week=${weekToUpdate}`);
        dataToProcess = JSON.parse(nextWeekResponse.getContentText());
      }
    }

    // --- Filter Games by Day based on user's 'config' property ---
    let originalEventCount = dataToProcess.events.length;
    let daySettings = null;
    try {
      const configString = PropertiesService.getUserProperties().getProperty('config');
      if (configString) {
        const config = JSON.parse(configString);
        if (config && config.matchupCustomization && config.matchupCustomization.days) {
          daySettings = config.matchupCustomization.days;
          Logger.log(`🔎 Found day settings in config: ${JSON.stringify(daySettings)}`);
        }
      }
    } catch (e) {
      Logger.log(`⚠️ Could not parse 'config' user property. Proceeding without filtering | Error: ${e.stack}`);
      daySettings = null;
    }

    if (daySettings) {
      const dayMap = { 0: 'includeSun', 1: 'includeMon', 2: 'includeTue', 3: 'includeWed', 4: 'includeThu', 5: 'includeFri', 6: 'includeSat' };
      dataToProcess.events = dataToProcess.events.filter(event => {
        const gameDay = new Date(event.date).getDay();
        const settingKey = dayMap[gameDay];
        return daySettings.hasOwnProperty(settingKey) && daySettings[settingKey] === true;
      });
      Logger.log(`📅 Filtered games from ${originalEventCount} to ${dataToProcess.events.length} based on config.`);
    }

    // --- Safety Check and User Confirmation Logic ---
    const allRelevantGamesPending = dataToProcess.events && dataToProcess.events.length > 0 &&
                                    dataToProcess.events.every(event => event.status.type.state === 'pre');
    
    let proceed = false;
    const spreadsAvailable = dataToProcess.events.some(event => event.competitions[0].odds);
    
    if (headless) {
      proceed = spreadsAvailable && allRelevantGamesPending;
      if (!allRelevantGamesPending) Logger.log(`⚠️ HEADLESS ABORT: Not all games in Week ${weekToUpdate} are pending.`);
    } else {
      if (!spreadsAvailable) {
        const nextWeekResponse = ui.alert('⭕ No Data', `Spread data is not yet available for Week ${weekToUpdate}.\n\nDo you want to check for next week's data (Week ${weekToUpdate + 1})?`, ui.ButtonSet.YES_NO);
        if (nextWeekResponse === ui.Button.YES) {
          fetchLatestSpreadsForWeek(false, weekToUpdate + 1);
        }
        return;
      }

      if (allRelevantGamesPending) {
        const response = ui.alert('❔ Confirm Update', `Spread data is available for ${dataToProcess.events.length} games in Week ${weekToUpdate}.\n\nDo you want to import/update this data?`, ui.ButtonSet.YES_NO);
        if (response === ui.Button.YES) {
          proceed = true;
        } else if (response === ui.Button.NO) {
          const checkNextWeek = ui.alert('Check Next Week?', `Would you like to check for spread data in Week ${weekToUpdate + 1} instead?`, ui.ButtonSet.YES_NO);
          if (checkNextWeek === ui.Button.YES) {
            fetchLatestSpreadsForWeek(false, weekToUpdate + 1);
          }
          return;
        }
      } else {
        const response = ui.alert(
          '⚠️ Week in Progress!', 
          `Some games for Week ${weekToUpdate} have already started or are complete. Spreads may be outdated.\n\nForce an update with available data anyway?`, 
          ui.ButtonSet.YES_NO_CANCEL
        );
        if (response === ui.Button.YES) {
          proceed = true;
        } else if (response === ui.Button.NO) {
          const checkNextWeek = ui.alert('Check Next Week?', `Would you like to check for spread data in Week ${weekToUpdate + 1} instead?`, ui.ButtonSet.YES_NO);
          if (checkNextWeek === ui.Button.YES) {
            fetchLatestSpreadsForWeek(false, weekToUpdate + 1);
          }
          return;
        }
      }
    }

    // --- Perform Update if Proceed is True ---
    if (proceed) {
       if (ui) ss.toast(`Updating spreads for Week ${weekToUpdate}...`, '🔄 IN PROGRESS');
      
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const columnIndices = {
        week: headers.indexOf('week'), awayTeam: headers.indexOf('awayTeam'), homeTeam: headers.indexOf('homeTeam'),
        overUnder: headers.indexOf('overUnder'), spread: headers.indexOf('spread'),
        spreadAutoFetched: headers.indexOf('spreadAutoFetched'), timeFetched: headers.indexOf('timeFetched')
      };
      
      const allData = sheet.getDataRange().getValues();
      let updatesMade = 0;
      const timeFetched = new Date();

      dataToProcess.events.forEach(event => {
        if (event.competitions[0].odds && event.competitions[0].odds.length > 0) {
          const apiAwayTeam = event.competitions[0].competitors.find(c => c.homeAway === 'away').team.abbreviation;
          const apiHomeTeam = event.competitions[0].competitors.find(c => c.homeAway === 'home').team.abbreviation;

          for (let i = 1; i < allData.length; i++) {
            if (allData[i][columnIndices.week] == weekToUpdate && allData[i][columnIndices.awayTeam] === apiAwayTeam && allData[i][columnIndices.homeTeam] === apiHomeTeam) {
              const odds = event.competitions[0].odds[0];
              sheet.getRange(i + 1, columnIndices.overUnder + 1).setValue(odds.overUnder);
              sheet.getRange(i + 1, columnIndices.spread + 1).setValue(odds.details);
              sheet.getRange(i + 1, columnIndices.spreadAutoFetched + 1).setValue(headless ? 1 : 0);
              sheet.getRange(i + 1, columnIndices.timeFetched + 1).setValue(timeFetched);
              updatesMade++;
              break;
            }
          }
        }
      });

      SpreadsheetApp.flush();
      const successMsg = `Spread data updated for ${updatesMade} games in Week ${weekToUpdate}.`;
      Logger.log(`✅ ${successMsg}`);
      if (ui) ss.toast(successMsg, `✅ UPDATED SPREADS`, 60);
    } else if (!headless) {
      const cancelMsg = `Update for Week ${weekToUpdate} cancelled or no data available.`;
      Logger.log(`🚫 ${cancelMsg}`);
      if (ui) ss.toast(cancelMsg, `🚫 CANCELED`, 60);
    }

  } catch (err) {
    Logger.log(`❌ Error in fetchLatestSpreadsForWeek | ${err.stack}`);
    if (ui) ui.alert('An error occurred. Please check the logs for details.');
  }
}


/**
 * This is the function the trigger will actually call.
 * It provides the correct parameters to your main fetchSchedule function.
 */
function runWeeklyFetch() {
  Logger.log(`🔄 Weekly auto-fetch trigger is running...`);
  fetchLatestSpreadsForWeek(true); // Run in headless mode to avoid prompts, delays  
}



// NFL GAMES - output by week input and in array format: [date,day,hour,minute,dayName,awayTeam,homeTeam,awayTeamLocation,awayTeamName,homeTeamLocation,homeTeamName]
function fetchGames(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (week == null) {
    week = fetchWeek();
  }
  try {
    const nfl = ss.getRangeByName(LEAGUE).getValues();
    let games = [];
    for (let a = 0; a < nfl.length; a++) {
      if (nfl[a][0] == week) {
        games.push(nfl[a].slice(1));
      }
    }
    return games;
  }
  catch (err) {
    let text = `Attempted to fetch ${LEAGUE} matches for week ${week} but no ${LEAGUE} data exists, fetching now...`;
    Logger.log(text);
    ss.toast(text);
    fetchSchedule(ss,null,week);
    return fetchGames(week);
  }
}

// NFL Schedule from ESPN API Scoreboard
function fetchMatchups() {
  let data = [];
  const obj = espnFetchJson(SCOREBOARD);
  let week = obj.season === 2 ? obj.week.number : (obj.season.type === 3 ? obj.week.number + REGULAR_SEASON : null);
  if (week === null) {
    throw new Error('Issue with the ESPN API for week');
  }
  for (let event = 0; event < obj.events.length; event++) {
    const date = new Date(obj.events[event].date);
    const hour = date.getHours();
    const minute = date.getMinutes();
    const day = date.getDay();
    const away = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'away')[0].team;
    const home = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'home')[0].team;
    const divisional = LEAGUE_DATA[home.abbreviation].division_opponents.indexOf(away.abbreviation) > -1 ? 1 : 0;
    const divName = divisional == 1 ? LEAGUE_DATA[home.abbreviation].division : '';
    const overUnder = (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].overUnder : '';
    const favorite = (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].details : '';

        // let bets = data[a][key][0].odds[0];
        // // Logger.log(data[a].shortName + ' - ' + JSON.stringify(bets));
        // arr.push([data[a].week.number,a+1,data[a].shortName.replace(" ","").replace(" ",""),
        // bets.overUnder,
        // bets.awayTeamOdds.team.abbreviation,(bets.awayTeamOdds.favorite ? parseFloat(-Math.abs(bets.spread)) : parseFloat(Math.abs(bets.spread))),
        // parseFloat(((parseFloat(bets.overUnder) + (bets.awayTeamOdds.favorite ? parseFloat(Math.abs(bets.spread)) : parseFloat(-Math.abs(bets.spread))))/2).toFixed(2)),           
        // bets.homeTeamOdds.team.abbreviation,(bets.homeTeamOdds.favorite ? parseFloat(-Math.abs(bets.spread)) : parseFloat(Math.abs(bets.spread))),
        // parseFloat(((parseFloat(bets.overUnder) + (bets.homeTeamOdds.favorite ? parseFloat(Math.abs(bets.spread)) : parseFloat(-Math.abs(bets.spread))))/2).toFixed(2))])


    let arr = [
      week,
      date,
      DAY[day].index,
      hour,
      minute,
      DAY[day].name,
      away.abbreviation,
      home.abbreviation,
      away.location,
      away.name,
      home.location,
      home.name,
      WEEKNAME.hasOwnProperty(week) ? WEEKNAME[week].name : 'Regular Season',
      divisional,
      divName,
      overUnder,
      favorite
    ];
    data.push(arr);
  }
  return data;
}

// LEAGUE LOGOS - Saves URLs to logos to a Document Property variable named "logos" {CURRENTLY UNUSED}
function fetchLogos(){
  let obj = {};
  let logos = {};
  try{
    obj = espnFetchJson(SCOREBOARD);
  }
  catch (err) {
    Logger.log(`⚠️ Fetch Logo error: ${err.stack}`);
    ui.alert(`⚠️ ESPN API ISSUE`,`The API for fetching logos isn't responding currently, try again in a moment.`,ui.ButtonSet.OK);
    throw new Error('⚠️ ESPN API issue, try later');
  }
  
  if (Object.keys(obj).length > 0) {
    let games = obj.events;
    // Loop through games provided and creates an array for placing
    for (let a = 0; a < games.length; a++){
      let competitors = games[a].competitions[0].competitors;
      let teamOne = competitors[0].team.abbreviation;
      let teamTwo = competitors[1].team.abbreviation;
      let teamOneLogo = competitors[0].team.logo;
      let teamTwoLogo = competitors[1].team.logo;
      logos[teamOne] = teamOneLogo;
      logos[teamTwo] = teamTwoLogo;
    }
    const docProps = PropertiesService.getDocumentProperties();
    try {
      let logoProp = docProps.getProperty('logos');
      let tempObj = JSON.parse(logoProp);
      if (Object.keys(tempObj).length < nflTeams) {
        docProps.setProperty('logos',JSON.stringify(logos));
      }
    }
    catch (err) {
      Logger.log(`⚠️ Error fetching logo object, creating one now`);
      docProps.setProperty('logos',JSON.stringify(logos));
    }
  }
  return logos;
}

// NFL OUTCOMES FUNCTIONS

/**
 * Creates and displays the custom HTML dialog for importing game scores.
 * This should be attached to a menu item.
 * The main function to launch the automated outcome import process.
 * This should be called from a menu item like "Import Scores from Live Data".
 */
function launchApiOutcomeImport() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('scoreImport')
      .setTitle(`${LEAGUE} Outcome Fetch`) 
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    // DEPRECATION NOTICE
    const ui = SpreadsheetApp.getUi();
    Logger.log(`⚠️ Could not launch new outcome import screen, user likely doesn't have "scoreImport.html" file available: ${err.stack}`);
    let newImportToolAlertMessage = ui.alert(`⚠️ DEPRECATION NOTICE`, `🆕 There's a new score import tool that can import previous week outcomes and scores.\n\n⭐ To update, please create a "scoreImport.html" using the content from either the template document or the GitHub project page, then try running again.`, ui.ButtonSet.OK_CANCEL);
  }
}

/**
 * Parses the events array from the API, categorizes games, and
 * extracts relevant outcome data.
 * @param {Array} apiEvents - The 'events' array from the API response.
 * @param {Object} gamePlan - The authoritative gamePlan for the week.
 * @returns {Object} An object containing categorized game outcomes.
 */
function parseApiEvents(apiEvents, gamePlan) {
  const statusMap = {
    'STATUS_SCHEDULED':'pregame',
    'STATUS_IN_PROGRESS':'active',
    'STATUS_POSTPONED':'postponed',
    'STATUS_FINAL':'complete'
    }
  const analysis = {
    pregame: [],
    active: [],
    postponed: [],
    complete: [],
    unknown: []
  };
  const gamePlanMatchups = new Set(gamePlan.games.map(g => `${g.awayTeam} @ ${g.homeTeam}`));

  apiEvents.forEach(event => {
    const eventName = event.shortName.replace(`VS`,`@`); // Case of neutral site a VS is used
    if (gamePlanMatchups.has(eventName)) {
      const status = statusMap[event.status.type.name];
      const homeTeam = event.competitions[0].competitors.find(c => c.homeAway === 'home');
      const awayTeam = event.competitions[0].competitors.find(c => c.homeAway === 'away');

      const gameData = {
        shortName: eventName,
        homeScore: parseInt(homeTeam.score, 10),
        awayScore: parseInt(awayTeam.score, 10)
      };

      if (status === 'complete') {
        // Game is finished, determine winner and margin
        gameData.winner = (gameData.homeScore > gameData.awayScore) ? homeTeam.team.abbreviation : awayTeam.team.abbreviation;
        gameData.loser = (gameData.homeScore < gameData.awayScore) ? homeTeam.team.abbreviation : awayTeam.team.abbreviation;
        if (gameData.homeScore === gameData.awayScore) gameData.winner = 'TIE';
        if (gameData.homeScore === gameData.awayScore) gameData.loser = 'TIE';
        gameData.margin = Math.abs(gameData.homeScore - gameData.awayScore);
        analysis.complete.push(gameData);
      } else if (status === 'postponed') {
        analysis.postponed.push(gameData);
      } else if (status === 'active') {
        analysis.active.push(gameData);
      } else if (status === 'pregame') {
        analysis.pregame.push(gameData);
      } else {
        analysis.unknown.push(gameData);
      }
    }
  });

  return analysis;
}

/**
 * Tool for fetching historic outcomes if new API week of scoreboard has advanced, may replace SCOREBOARD fetching
 * Parses the new proTeams array from the API, categorizes games based on
 * percentComplete, and filters/extracts relevant outcome data by cross-referencing
 * against ALL gamePlans in formsData.
 *
 * @param {Array} apiProTeams - The 'proTeams' array from the new API response.
 * @param {Object} formsData - The object containing all authoritative gamePlans keyed by week number.
 * @returns {Object} An object containing categorized game outcomes.
 */
function parseAllApiEvents(apiProTeams, formsData) {
  let docProps;
  if (!formsData) docProps = PropertiesService.getDocumentProperties();

  formsData = formsData || JSON.parse(docProps.getProperty('forms'));
  
  apiProTeams = apiProTeams || espnFetchJson(`${schedulePrefix}${fetchYear()}${scheduleSuffix}`).settings.proTeams || [];

  // 1. Create a map for Team ID to Abbreviation (e.g., {11: "Ind"})
  const teamIdMap = {};
  apiProTeams.forEach(team => {
    if (team.abbrev != 'FA') teamIdMap[team.id] = team.abbrev.toUpperCase();
  });
  
  const getStatus = (game) => {
    if (game.percentComplete === 0) return 'pregame';
    if (game.percentComplete === 100) return 'complete';
    if (game.percentComplete > 0) return 'active';
    if (game.postponed || (game.postponed == 'true') && game.percentComplete == 0) return 'postponed';
    Logger.log(JSON.stringify(game));
    Logger.log(game);
    return 'unknown';
  };
  
  const analysis = {};
  // 2. Collect all valid matchups from ALL game plans (all weeks) for filtering
  const allGamePlanMatchups = new Set();
  // Iterate through all week properties in formsData
  Object.values(formsData).forEach(formData => {
    if (formData && formData.gamePlan && formData.gamePlan.games) {
      formData.gamePlan.games.forEach(g => {
        // The expected format is "AWAY_ABBREV @ HOME_ABBREV"
        allGamePlanMatchups.add(`${g.awayTeam} @ ${g.homeTeam}`);
      });
    }
  });

  // Keep track of games already processed to avoid duplicates
  const processedGameIds = new Set();

  // 3. Iterate through all games across all teams and weeks
  apiProTeams.forEach(team => {
    const proGamesByPeriod = team.proGamesByScoringPeriod;
    
    // Iterate over each scoring period (week)
    Object.keys(proGamesByPeriod).forEach(scoringPeriodId => {
      // scoringPeriodId is a string, e.g., "1", "2"
      
      const games = proGamesByPeriod[scoringPeriodId];
      
      // Iterate over each game in the period (usually one)
      games.forEach(game => {
        
        // Skip if we've already processed this unique game ID
        if (processedGameIds.has(game.id)) return;
        processedGameIds.add(game.id);

        const homeAbbrev = teamIdMap[game.homeProTeamId];
        const awayAbbrev = teamIdMap[game.awayProTeamId];

        // If team ID resolution fails, skip (shouldn't happen with valid data)
        if (!homeAbbrev || !awayAbbrev) return; 

        // Construct the event name for cross-referencing
        const eventName = `${awayAbbrev} @ ${homeAbbrev}`;

        // 4. Cross-reference against the collective game plan
        if (allGamePlanMatchups.has(eventName)) {
          if (!analysis[scoringPeriodId]) {
            analysis[scoringPeriodId] = {
              pregame: [],
              active: [],
              postponed: [],
              complete: [],
              unknown: []
            };
          }

          const status = getStatus(game);

          const gameData = {
            date: game.date,
            shortName: eventName,
            homeScore: game.homeScore,
            awayScore: game.awayScore,
            winner: null
          };

          if (status === 'complete') {
            gameData.winner = (gameData.homeScore > gameData.awayScore) ? homeAbbrev : awayAbbrev;
            gameData.loser = (gameData.homeScore < gameData.awayScore) ? homeAbbrev : awayAbbrev;
            if (gameData.homeScore === gameData.awayScore) {
                gameData.winner = 'TIE';
                gameData.loser = 'TIE';
            }
            gameData.margin = Math.abs(gameData.homeScore - gameData.awayScore);
            if (scoringPeriodId == 5) {
              Logger.log(JSON.stringify(gameData));
            }
          }

          if (analysis[scoringPeriodId][status]) {
            analysis[scoringPeriodId][status].push(gameData);
          } else {
            analysis[scoringPeriodId]['unknown'].push(gameData); // Fallback for safety
          }

        }
      });
    });
  });

  Object.keys(analysis).forEach(scoringPeriodId => {
    Object.keys(analysis[scoringPeriodId]).forEach(status => {
      analysis[scoringPeriodId][status].sort((a, b) => a.date - b.date);
    });
  });

  return analysis;
}

function getScoreImportData() {
  const apiProTeams = espnFetchJson(`${schedulePrefix}${fetchYear()}${scheduleSuffix}`).settings.proTeams || [];
  if (!apiProTeams) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error fetching API data, canceling...`,`API ERROR`);
    throw new Error("Could not load API data.");
  }
  const docProps = PropertiesService.getDocumentProperties();
  const formsData = JSON.parse(docProps.getProperty('forms'));
  const tiebreakerInclude = JSON.parse(docProps.getProperty('configuration')).tiebreakerInclude;
  const analysis = parseAllApiEvents(apiProTeams, formsData)
  const apiWeek = fetchWeek(null, true);
  return {
    analysis: analysis,
    formsData: formsData,
    week: apiWeek,
    tiebreakerInclude: tiebreakerInclude,
    leagueData: LEAGUE_DATA
  }
}


/**
 * Takes the parsed API outcomes and writes them to the
 * user-facing weekly sheet and the master NFL_OUTCOMES sheet.
 */
function updateSheetsWithApiOutcomes(ss, week, completedGames, formsData, booleanOutput) {
  ss = ss || fetchSpreadsheet();
  if (completedGames.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('⭕ No completed games to import.');
    return false;
  }
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  
  formsData = formsData || JSON.parse(docProps.getProperty('forms')) || {};
  if (!formsData || formsData.hasOwnProperty('games')) {
    formsData = JSON.parse(docProps.getProperty('forms'));
  }
  
  let errorMessage = '';

  const weeklySheet = ss.getSheetByName(`${weeklySheetPrefix}${week}`);
  if (weeklySheet) {
    const weeklySheetMap = outcomeDataValidationMapping(week, formsData, `${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const weeklySheetWinnersRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const weeklySheetMarginsRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`);
    let weeklySheetWinners = weeklySheetWinnersRange.getValues();
    let weeklySheetMargins = weeklySheetMarginsRange.getValues();
    if (weeklySheetMap) {
      completedGames.forEach(game => {
        // Returns a map of games in the form {'AWAY @ HOME:colIndex} -- Row is 1-based for sheets placement
        const colIndex = weeklySheetMap[game.shortName]; // Allows for the first row of map data to be first row of named range
        if (colIndex && weeklySheetWinners[0][colIndex-1] == '') {
          weeklySheetWinners[0][colIndex-1] = game.winner;
        }
        if (colIndex && weeklySheetMargins[0][colIndex-1] == '') {
          weeklySheetMargins[0][colIndex-1] = game.margin;
        }
      });
      try {
        Logger.log(`🔄 Placing values for the winners and margins of ${completedGames.length} game(s) now.`);
        weeklySheetWinnersRange.setValues(weeklySheetWinners);
        weeklySheetMarginsRange.setValues(weeklySheetMargins);
      } catch(err) {
        try {
          Logger.log(`⚠️ Error placing values for the weekly from games for week ${week}. Attempting to place them based on gamePlan order...`)
          const horizontalWinners = weeklySheetWinners[0].map((value, index) => {
            if (value == '' && completedGames[index]) {
              return completedGames[index].winner;
            } else {
              return weeklySheetWinners[0][index];
            }
          });
          const horizontalMargins = weeklySheetMargins[0].map((value, index) => {
            if (value == '' && completedGames[index]) {
              return completedGames[index].margin;
            } else {
              return weeklySheetMargins[0][index];
            }
          });
          weeklySheetWinnersRange.setValues([horizontalWinners]);
          weeklySheetMarginsRange.setValues([horizontalMargins]);
          Logger.log(`✅ Backup outcome recording worked for ${completedGames.length} game(s)`);
        } catch(err) {
          Logger.log(`❗ Unable to place any outcomes. Manual mode recommended or rebuild WEEKLY sheet | ERROR: ${err.stack}`);
          errorMessage += `Attempted to bring in pick 'ems sheet weekly outcomes and failed to place outcomes even with backup solution.`
        }
      }
    }
    if (config.tiebreakerInclude) {
      const weekGamePlan = formsData[week].gamePlan;
      // Use the game the form actually asked about. Weeks built before the tiebreaker
      // game was pinned into the plan fall back to deriving it the same way.
      const tiebreakerGame = weekGamePlan.tiebreakerGame || getTiebreakerGame(weekGamePlan.games);
      const tiebreakerMatchup = tiebreakerGame ? `${tiebreakerGame.awayTeam} @ ${tiebreakerGame.homeTeam}` : null;

      const placeOutcome = (outcomeRangeName, groupRangeName, value, label) => {
        let range = ss.getRangeByName(outcomeRangeName);
        if (!range) {
          // Fall back to the outcome cell three rows below the member entry block
          const groupRange = ss.getRangeByName(groupRangeName);
          if (groupRange) range = groupRange.getSheet().getRange(groupRange.getLastRow() + 3, groupRange.getColumn());
        }
        if (range) {
          range.setValue(value);
          Logger.log(`👔 Placed ${label} of ${value} from ${tiebreakerMatchup} (week ${week}).`);
          return true;
        }
        Logger.log(`❗👔 Found ${label} of ${value} for ${tiebreakerMatchup} but no ${outcomeRangeName} range to place it in`);
        ss.toast(`Found ${label} of ${value} for ${tiebreakerMatchup} but could not place it`,`❗ TIEBREAKER NOT PLACED`);
        return false;
      };

      if (!tiebreakerMatchup) {
        Logger.log(`⚠️ No tiebreaker game could be determined for week ${week}; skipping tiebreaker outcomes.`);
      } else {
        Logger.log(`⚖️ Week ${week} tiebreaker game is ${tiebreakerMatchup}. Checking for outcome availability...`);
        const tiebreakerMatchupDetails = completedGames.find(game => game.shortName === tiebreakerMatchup);
        if (tiebreakerMatchupDetails) {
          const awayScore = parseInt(tiebreakerMatchupDetails.awayScore);
          const homeScore = parseInt(tiebreakerMatchupDetails.homeScore);
          const combinedScore = awayScore + homeScore;
          // TB2 is the winning team's points; on a tie both teams scored the same
          const winningScore = Math.max(awayScore, homeScore);
          Logger.log(`🔥 ${tiebreakerMatchupDetails.winner} won the matchup — combined score ${combinedScore}, winning team scored ${winningScore}`);
          if (Number.isFinite(combinedScore) && Number.isFinite(winningScore)) {
            placeOutcome(`${LEAGUE}_TIEBREAKER_${week}_OUTCOME`, `${LEAGUE}_TIEBREAKER_${week}`, combinedScore, 'combined score');
            placeOutcome(`${LEAGUE}_TIEBREAKER2_${week}_OUTCOME`, `${LEAGUE}_TIEBREAKER2_${week}`, winningScore, 'winning team score');
          } else {
            Logger.log(`❗👔 Scores for ${tiebreakerMatchup} were not numeric (away: ${tiebreakerMatchupDetails.awayScore}, home: ${tiebreakerMatchupDetails.homeScore}); skipping placement.`);
          }
        } else {
          Logger.log(`⏩ Tiebreaker matchup for week ${week} of ${tiebreakerMatchup} incomplete, skipping tiebreaker for now.`)
          ss.toast(`Tiebreaker matchup for week ${week} of ${tiebreakerMatchup} incomplete, skipping tiebreaker for now.`,`⏩ TIEBREAKER NOT AVAILABLE`);
        }
      }
    } else {
      Logger.log(`👔 No tiebreaker configured for the pool.`);
    }
  }

  const outcomesSheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);
  if (outcomesSheet) {
    const outcomesSheetMap = outcomeDataValidationMapping(week, formsData, `${LEAGUE}_OUTCOMES_${week}`);
    // Find the corresponding row in the OUTCOMES sheet
    const outcomesSheetWinnersRange = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
    const outcomesSheetMarginsRange = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);
    let outcomesSheetWinners = outcomesSheetWinnersRange.getValues();
    let outcomesSheetMargins = outcomesSheetMarginsRange.getValues();
    if (outcomesSheetMap) {
      completedGames.forEach(game => {
        // Returns a map of games in the form {'AWAY @ HOME:rowIndex} -- Row is 1-based for sheets placement
        const rowIndex = outcomesSheetMap[game.shortName]; // Allows for the first row of map data to be first row of named range
        if (rowIndex && outcomesSheetWinners[rowIndex-1][0] == '') {
          outcomesSheetWinners[rowIndex-1][0] = game.winner;
        }
        if (rowIndex && outcomesSheetMargins[rowIndex-1][0] == '') {
          outcomesSheetMargins[rowIndex-1][0] = game.margin;
        }
      });
      try {
        Logger.log(`🔄 Placing values for the winners and margins of ${completedGames.length} game(s) now.`);
        outcomesSheetWinnersRange.setValues(outcomesSheetWinners);
        outcomesSheetMarginsRange.setValues(outcomesSheetMargins);
      } catch(err) {
        try {
          Logger.log(`⚠️ Error placing values for the outcomes from games for week ${week}. Attempting to place them based on gamePlan order...`)
          const verticalWinners = outcomesSheetWinners.map((value, index) => {
            if (value[0] == '' && completedGames[index]) {
              return [completedGames[index].winner];
            } else {
              return [outcomesSheetWinners[index][0]];
            }
          });
          const verticalMargins = outcomesSheetMargins.map((value, index) => {
            if (value[0] == '' && completedGames[index]) {
              return [completedGames[index].margin];
            } else {
              return [outcomesSheetMargins[index][0]];
            }
          });
          outcomesSheetWinnersRange.setValues(verticalWinners);
          outcomesSheetMarginsRange.setValues(verticalMargins);
          Logger.log(`✅ Backup outcome recording worked for ${completedGames.length} game(s)`);
        } catch(err) {
          Logger.log(`❗ Unable to place any outcomes. Manual mode recommended or rebuild OUTCOMES sheet | ERROR: ${err.stack}`);
          errorMessage += `\n\nAttempted to bring in ${LEAGUE}_OUTCOMES sheet weekly outcomes and failed to place outcomes even with backup solution.`
        }
      }
    }
  }
  if (errorMessage.length > 0) {
    if (booleanOutput) {
      const ui = fetchUi();
      ui.alert(`⚠️ Outcome Import Issue`, `${errorMessage}\n\nTry again later or reach out for support.\n\nPicks ${(config.pickemsInclude && config.pickemsAts) || (config.survivorInclude && config.survivorAts) || (config.eliminatorInclude && config.eliminatorAts) ? 'and margins ' : ''} can always be manually entered.`, ui.ButtonSet.OK);  
      return false; // Used for deprecated version
    }
    return {"message":`⚠️ Outcome Import Issue: ${errorMessage}\n\nTry again later or reach out for support.\n\nPicks ${(config.pickemsInclude && config.pickemsAts) || (config.survivorInclude && config.survivorAts) || (config.eliminatorInclude && config.eliminatorAts) ? 'and margins ' : ''} can always be manually entered.`}
  } else {
    if (booleanOutput) return true; // Used for deprecated version
    return {"message":`✅ SUCCESS!\n\nImported outcomes for ${completedGames.length} complete games for week ${week}.`}
  }
}

/** 
 * Function to receive infor that is used to process the existing validation in the OUTCOMES sheet and provide a map for placing values.
 */
function outcomeDataValidationMapping(week, formsData, namedRangeName) {
  try {
    namedRangeName = namedRangeName || `${LEAGUE}_OUTCOMES_${week}`;
    const namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRangeName);
    
    if (!namedRange) {
      Logger.log(`❗ Named range "${namedRangeName}" not found`);
      return false;
    }
    const numRows = namedRange.getNumRows();
    const numCols = namedRange.getNumColumns();
    
    Logger.log(`🔍 Checking named range: ${namedRangeName} (${numRows}x${numCols} cells)`);
    const games = formsData[week]?.gamePlan?.games;
    if (!games) {
      Logger.log(`❌ No games found in formsData[${week}].gamePlan.games`);
      return false;
    }
    const gamesArray = games.map(contest => {
      return `${contest.awayTeam} @ ${contest.homeTeam}`;
    })
    
    Logger.log(`📋 These matchup exist for week ${week}: ${gamesArray}`);    
    let found = false;
    let matchupMap = {};
    // Iterate through each cell in the named range
    const horizontal = numRows == 1 ? true : false;
    for (let row = 1; row <= numRows; row++) {
      for (let col = 1; col <= numCols; col++) {
        const cell = namedRange.getCell(row, col);
        
        // Get data validation for this cell
        const validation = cell.getDataValidation();
        
        if (validation) {
          const criteria = validation.getCriteriaValues();
          
          if (criteria && criteria.length >= 2) {
            const firstOption = criteria[0][0];
            const secondOption = criteria[0][1];
            
            const combinedMatchup = `${firstOption} @ ${secondOption}`;
            if (!gamesArray.includes(combinedMatchup)) {
              Logger.log(`❌ Matchup "${combinedMatchup}" not found in games object`);
              allValid = false;
            } else {
              matchupMap[combinedMatchup] = horizontal ? col : row;
              if (!found) found = true;
              // Logger.log(`✅ Matchup "${combinedMatchup}" found in games object in cell (${row},${col})`);
            }
          } else {
            Logger.log(`⚠️ Cell (${row},${col}): Data validation found but insufficient criteria`);
            allValid = false;
          }
        } else {
          Logger.log(`⛔ Cell (${row},${col}): No data validation found`);
        }
      }
    }    
    if (found) {
      Logger.log(`✅ Map created with these values: ${JSON.stringify(matchupMap)}`);
      Logger.log(`↩️ Returning matchup map for placing game outcomes within the OUTCOMES sheet`);
      return matchupMap;
    } else {
      Logger.log(`🚫 Unable to identify any named ranges within the provided week for the OUTCOME sheet`);
      return false;
    }
    
  } catch (err) {
    Logger.log(`⚠️ Error validating named range: ${err.stack}`);
    return false;
  }
}


// ============================================================================================================================================
// FORM FUNCTIONS
// ============================================================================================================================================

/**
 * [CORRECTED & BULLETPROOF] Gathers and aggregates all data for the dashboard.
 * This version is resilient to missing or incomplete data properties.
 */
function getFormDashboardData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    
    // --- Provide robust defaults for all major properties ---
    const config = JSON.parse(docProps.getProperty('configuration')) || {};
    const memberData = JSON.parse(docProps.getProperty('members')) || { "memberOrder": [], "members": {} };
    const formsObject = JSON.parse(docProps.getProperty('forms')) || {};
    
    const apiWeek = fetchWeek(null, true);
    
    let formsList = [];
    
    // 1. Find all form data from the forms object
    for (const week in formsObject) {
      // Use optional chaining (?.) for safe property access
      const form = formsObject[week];
      if (!form) continue; // Skip if the form entry is null for some reason

      // Create a new object to avoid modifying the original
      const formDetails = { 
        week: parseInt(week),
        ...form
      };
      // 2. Augment the data, using safe fallbacks for every value
      try {
        formDetails.isActive = FormApp.openById(formDetails.formId).isAcceptingResponses();
      } catch (err) {
        formDetails.isActive = false;
      }
      const gamePlan = formDetails.gamePlan;
      formDetails.respondents = form.respondents;
      formDetails.responseCount = form.respondents.length;
      formDetails.nonRespondents = form.nonRespondents;

      formDetails.membershipLocked = gamePlan?.membershipLocked || false;
      formDetails.newMembers = form.newMembers;
      formDetails.gameCount = gamePlan?.games?.length || 0;
      
      formDetails.pickemsInclude = gamePlan?.pickemsInclude || false;
      formDetails.pickemsAts = gamePlan?.pickemsAts || false;
      formDetails.survivorInclude = gamePlan?.survivorInclude  || false;
      formDetails.survivorAts = gamePlan?.survivorAts  || false;
      formDetails.eliminatorInclude = gamePlan?.eliminatorInclude  || false;
      formDetails.eliminatorAts = gamePlan?.eliminatorAts  || false;;
      formsList.push(formDetails);
    }
    
    formsList.sort((a, b) => a.week - b.week);
    
    // --- Calculation logic with safe defaults ---
    const totalMembers = memberData.memberOrder?.length || 0;
    const totalResponses = formsList.reduce((acc, f) => acc + f.responseCount, 0);
    const totalPossibleResponses = totalMembers * formsList.length;
    const templateId = docProps.getProperty('templateId');
    Logger.log(templateId);
    return {
      groupName: config.groupName || `${LEAGUE} Picks Pool`,
      forms: formsList,
      totalMembers: totalMembers,
      apiWeek: apiWeek,
      templateId: templateId, // Future use in formManager panel
      overallResponseRate: totalPossibleResponses > 0 ? (totalResponses / totalPossibleResponses) : 0      
    };
  } catch (err) {
    Logger.log(`⚠️ Error in getFormDashboardData: ${err.stack}`);
    throw new Error(`⚠️ Could not load form data. ${err.message}`);
  }
}

/**
 * The main function to launch the forms panel
 */
function launchFormManager() {
    const html = HtmlService.createHtmlOutputFromFile('formManager').setWidth(1000).setHeight(115);
    SpreadsheetApp.getUi().showModalDialog(html, 'Form Manager');
}

/**
 * Toggles the 'isAcceptingResponses' status of a Google Form
 * Also manages the onFormSubmit trigger for that form.
 * @param {string} formId The ID of the form to toggle.
 * @returns {Object} An object containing the new active status.
 */
function toggleFormStatus(formId) {
  try {
    const form = FormApp.openById(formId);
    const currentState = form.isAcceptingResponses();
    const newState = !currentState;
    form.setAcceptingResponses(newState);

    setFormSubmitTrigger(formId, newState);
    
    return { success: true, newStatus: newState, newSyncStatus: newState };
  } catch (err) {
    Logger.log(`⚠️ Failed to toggle status for form ${formId}: ${err.stack}`);
    throw new Error(`⚠️ Could not update form status. ${err.message}`);
  }
  
}

/**
 * Enables or disables the onFormSubmit trigger for a specific form.
 * @param {string} formId The ID of the form to manage the trigger for.
 * @param {boolean} shouldBeEnabled The desired state for the trigger.
 * @returns {Object} A success message.
 */
function setFormSubmitTrigger(formId, shouldBeEnabled) {
  try {
    let found = false;
    // Clear out existing trigger(s) for form ID
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      if (trigger.getTriggerSourceId() === formId) {
        found = true;
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`❌ Trigger DELETED for form ID: ${formId}`);
      }
    });

    if (!found) {
      Logger.log(`⭕ Unable to locate an existing trigger for the form ID ${formId} among ${allTriggers.length} triggers.`)
    }

    let toastMessage = '';

    if (shouldBeEnabled) {
      // If enabling, create a new trigger
      const form = FormApp.openById(formId);
      ScriptApp.newTrigger('handleFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      toastTitle = `✅ TRIGGER ADDED`;
      toastMessage = `Auto-sync trigger has been created for the form.`;
      Logger.log(`✅ Auto-sync trigger ENABLED for form ID: ${formId}`);
    } else {
      // If disabling, notify user
      toastTitle = `❌ TRIGGER DELETED`;
      toastMessage = `Auto-sync trigger has been removed for the form.`;
    }

    // Store the preference
    const formsData = fetchProperties('forms');
    const week = getWeekFromFormId(formId);
    if (week && formsData[week]) {
      formsData[week].autoSync = shouldBeEnabled;
      saveProperties('forms', formsData);
    }
    
    // Display the toast message.
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMessage,toastTitle);

    return { success: true, newStatus: shouldBeEnabled };
  } catch (err) {
    Logger.log(`⚠️ Failed to set trigger for form ${formId}: ${err.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${err.message}`, '❌ FAILED', 10);
    throw new Error(`⚠️ Could not update trigger. ${err.message}`);
  }
}

/**
 * This is the function that the one-time trigger will execute.
 * It now looks up the form ID from properties using its own trigger ID.
 *
 * @param {Object} e The event object passed by the time-based trigger.
 */
function executeFormLock(e) {
  // Get the unique ID of the trigger that just ran
  const triggerId = e.triggerUid;
  if (!triggerId) {
    Logger.log(`⭕ executeFormLock ran but could not identify trigger ID.`);
    return;
  }
  
  // Look up the trigger metadata in Document Properties
  const docProps = PropertiesService.getDocumentProperties();
  const triggerMetaProperty = docProps.getProperty('triggerMeta_' + triggerId);
  
  if (!triggerMetaProperty) {
    Logger.log(`⭕ Could not find metadata for trigger ID ${triggerId}. Aborting lock.`);
    // Attempt to clean up the trigger anyway
    deleteTriggerById(triggerId);
    return;
  }
  
  const metadata = JSON.parse(triggerMetaProperty);
  const formId = metadata.formId;

  try {
    Logger.log(`▶️ Executing one-time lock for form ID: ${formId}`);
    
    // Lock the form
    FormApp.openById(formId).setAcceptingResponses(false);
    
    // Delete the trigger and its metadata
    deleteTriggerById(triggerId);
    docProps.deleteProperty('triggerMeta_' + triggerId);

    Logger.log(`✅ Successfully executed and deleted one-time lock trigger for form ID: ${formId}`);
  } catch (err) {
    Logger.log(`⚠️ Failed to execute lock for form ID ${formId}. Error: ${err.toString()}`);
    // Attempt to clean up anyway.
    deleteTriggerById(triggerId);
    docProps.deleteProperty('triggerMeta_' + triggerId);
  }
}

/**
 * Creates a one-time trigger to lock a form and stores
 * metadata linking the trigger's ID to the form's ID.
 *
 * @param {string} formId The ID of the form to be locked.
 * @param {Object} gamePlan The gamePlan object for that week's form.
 */
function setOneTimeFormLockTrigger(formId, gamePlan) {
  if (!formId || !gamePlan || !gamePlan.games || gamePlan.games.length === 0) return;

  let earliestKickoff = null;
  gamePlan.games.forEach(game => {
    const gameTime = new Date(game.date);
    if (!earliestKickoff || gameTime < earliestKickoff) {
      earliestKickoff = gameTime;
    }
  });

  if (!earliestKickoff || earliestKickoff < new Date()) {
    Logger.log(`⚠️ Cannot set form lock trigger: earliest kickoff is in the past.`);
    return;
  }

  // Create the time-based trigger correctly.
  const trigger = ScriptApp.newTrigger('executeFormLock')
    .timeBased()
    .at(earliestKickoff)
    .create();

  // Get the unique ID of the new trigger
  const triggerId = trigger.getUniqueId();

  // Store metadata in properties, linking the trigger's ID to the form's ID.
  const metadata = { formId: formId, week: gamePlan.week };
  PropertiesService.getDocumentProperties().setProperty('triggerMeta_' + triggerId, JSON.stringify(metadata));

  Logger.log(`🔏 Scheduled one-time form lock for ${earliestKickoff.toLocaleString()} with trigger ID ${triggerId}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Form will automatically lock at first kickoff.`,`🔏 AUTO LOCK ENABLED`);
}

/**
 * A simple utility to delete a trigger by its unique ID.
 */
function deleteTriggerById(triggerId) {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getUniqueId() === triggerId) {
      ScriptApp.deleteTrigger(trigger);
      break;
    }
  }
}

/** 
 * A reverse-lookup to find a week by form ID.
*/
function getWeekFromFormId(formId) {
  const formsData = fetchProperties('forms');
  for (const week in formsData) {
    if (formsData[week].formId === formId) {
      return week;
    }
  }
  return null;
}

/**
 * onFormSubmit trigger will call this function
 * Wrapper that determines the week and calls our main sync function.
 * @param {Object} e: The event object passed by the onFormSubmit trigger.
 */
function handleFormSubmit(e) {
  try {
    // 1. Get the Form object and its unique ID from the event object.
    const form = e.source;
    const submittedFormId = form.getId();
    let week = null;

    // 2. Look up the form ID in our 'forms' property.
    const formsData = fetchProperties('forms'); // Your existing helper
    if (formsData && Object.keys(formsData).length > 0) {
      for (const weekNum in formsData) {
        if (formsData[weekNum].formId === submittedFormId) {
          week = parseInt(weekNum, 10);
          break; // We found our match, no need to loop further.
        }
      }
    }

    // 3. Fallback to parsing the title if the lookup fails (optional but safe).
    if (!week) {
      Logger.log(`🔎 Could not find form ID ${submittedFormId} in the 'forms' property. Falling back to parsing title.`);
      const formTitle = form.getTitle();
      const weekMatch = formTitle.match(/Week (\d+)/);
      if (weekMatch && weekMatch[1]) {
        week = parseInt(weekMatch[1], 10);
      }
    }

    // 4. If we have a week, run the main sync function.
    if (week) {
      Logger.log(`🔄 Form submit detected for Week ${week}. Running sync...`);
      // Run our main, robust sync function.
      syncFormResponses(week);
    } else {
      Logger.log(`⚠️ CRITICAL: Could not determine week for submitted form with ID: ${submittedFormId} and Title: "${form.getTitle()}". Sync aborted.`);
    }
  } catch (err) {
    // Log any errors that occur during the sync process itself.
    Logger.log(`⚠️ An error occurred during the onFormSubmit sync: ${err.stack}`);
  }
}

/**
 * Form creation process controller
 * Includes a preliminary check for the existence of the LEAGUE schedule sheet, prompts the user to fetch it if it's missing.
 */
function launchFormBuilder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // Get the spreadsheet object once
  const ui = SpreadsheetApp.getUi();
  const docProps = PropertiesService.getDocumentProperties();
  
  // --- Check 1: Configuration ---
  if (!docProps.getProperty('configuration')) {
    Logger.log(`No configuration data present, please begin by configuring the pool`);
    if (ui.alert(`⚠️ Configuration Missing`, `No configuration data found...`, ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
      launchConfiguration();
    } else {
      Logger.log(`⛔ Form creation canceled due to no configuration found.`);
      ss.toast(`Form creation canceled due to no configuration found.`,`⛔ NO SCHEDULE DATA`);
    }
    return;
  }
  
  // --- Check 2: Members ---
  const config = JSON.parse(docProps.getProperty('configuration'));
  let openEnrollment = false;
  if (!docProps.getProperty('members')) {
    Logger.log(`👻 No members found in your setup.`);
    ss.toast(`No members found in your setup.`,`👻 EMPTY MEMBERS`);
    let alertText = `No members data found. `;
    if (config.membershipLocked) {
      alertText += `Would you like to configure some initial members to be selectable via the dropdown on the form?\n\nYou have your pool membership locked, selecting "YES" will bring up the "Member Management" panel to enter initial members. Then restart the form builder when completed.\n\nIf you select "NO", the membership will be unlocked and a text entry box will be provided for the first week of the form for all members to join.`;
    } else {
      alertText += `Your membership is already set to open for joining via the form. However, would you like to enter any initial members?\n\nSelecting "YES" will bring up the "Member Management" panel to enter initial members. Then restart the form builder when completed.\n\nIf you select "NO", a text entry box will be provided for the first week of the form for all members to join.`;
    }
    let membersMissing = ui.alert(`⚠️ No Members`,`You haven't entered any members yet, would you like to do that now?\n\nSelect "Yes" to launch the member configuration panel, "No" to skip to form creation without any members (fully open enrollment)`, ui.ButtonSet.YES_NO_CANCEL);
    if (membersMissing === ui.Button.YES) {
      launchMemberPanel();
      return;
    } else if (membersMissing === ui.Button.CANCEL) {
      ss.toast(`Canceled form creation. Please check if you'd like to include new members and try again.`,`🛑 FORM BUILDER STOPPED`);
    } else {
      if (config.membershipLocked) {
        config.membershipLocked = false;
        saveProperties('configuration', config);
        ss.toast('Membership has been unlocked.', '🔓 MEMBERSHIP UNLOCKED');
        openEnrollment = true;
      } else {
        ss.toast('Cofirmed open enrollment for form (text field entry for all members in initial week), moving on.', '🔓 OPEN ENROLLMENT');
      }
    }
  }

  // --- [THE NEW LOGIC] Check 3: Schedule Data ---
  const scheduleSheet = ss.getSheetByName(LEAGUE); // e.g., 'NFL'
  if (!scheduleSheet) {
    const response = ui.alert(
      '⚠️ Schedule Data Missing',
      `The required '${LEAGUE}' schedule data sheet was not found. This is necessary to build the form matchups.\n\nWould you like to fetch and import the schedule data now?`,
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      try {
        // Run the fetchSchedule function with the specified parameters
        ss.toast(`Fetching ${LEAGUE} schedule, this may take a moment...`, '📡 SCHEDULE FETCH');
        // Let the function auto-detect the year and current week, set auto=false, overwrite=true
        fetchSchedule(ss, null, null, false, true); 
        ss.toast('Schedule data imported successfully!', '✅ SCHEDULE DATA IMPORTED');
        // After fetching, we can proceed.
      } catch (err) {
        ui.alert('Error', `Failed to fetch schedule data: ${err.message}`, ui.ButtonSet.OK);
        return; // Stop if the fetch fails
      }
    } else {
      // User declined to fetch the data
      Logger.log(`⛔ Form creation canceled, schedule data is required and declined to be brought in.`);
      ss.toast(`Form creation canceled, schedule data is required and declined to be brought in.`,`⛔ NO SCHEDULE DATA`);
      return;
    }
  }

  try {
    if (!checkFileExists(docProps.getProperty('templateId'))) {
      Logger.log(`📄 A template file needs to be created, routing to prompts...`);
      ss.toast(`A template file needs to be created, routing to prompts...`,`📄 TEMPLATE FILE NEEDED`);
      templateCreationPrompt(ss);
    } else {
      const htmlTemplate = HtmlService.createTemplateFromFile('formCreatorPanel');
      const htmlOutput = htmlTemplate.evaluate().setWidth(700).setHeight(210);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Create Form${openEnrollment ? ' - Open Enrollment' : ''}`);
    }
  } catch (err) {
    Logger.log(`⚠️ Error starting form creation: ${err.stack}`);
    ui.alert(`⚠️ FORM BUILD ERROR`,`An error occurred while launching the form builder:\n\n${err.message}`,ui.ButtonSet.OK);
  }
}

/**
 * Encapsulates the entire first-run and theme customization workflow.
 */
function templateCreationPrompt(ss,ui) {
  ui = ui || fetchUi();
  try {
    const templateForm = getTemplateForm();
    // Null check has to come first -- getTemplateForm() returns null when the user cancels,
    // and calling getId() on it threw a TypeError that the catch below then mislabeled.
    if (!templateForm) {
      Logger.log(`⛔ No template form available (user canceled or creation failed)`);
      return;
    }
    Logger.log(`📄 Template Form ${templateForm.getId()}`);

    let response = ui.alert(
      '🎨 Customize Form Theme (One Time Only)',
      `Each week's form will utilize this template file to generate a fresh form.\n\nBefore creating your first weekly form, would you like to customize the template?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      showLinkDialog(templateForm.getEditUrl(), '🎨 Template Customization', 'Form Template',`\nOnly modify the header image and theme colors of the form. Once modified, close the form and return here to restart the form builder tool.\n\nNote: changing the header image should automatically modify the color palette.`);
    }
  } catch (err) {
    if (err.message.includes("CANCELED_BY_USER")) {
      ss.toast('Form creation canceled by user when running a form building operation',`⛔ FORM CREATION CANCELED`);
      Logger.log(`⛔ Form creation canceled by user`);
    } else {
      ui.alert(`⚠️ FORM BUILD ERROR`,`An unexpected error occurred:\n\n${err.message}`,ui.ButtonSet.OK);
      Logger.log(`⚠️ Error occurred during form building process: ${err.stack}`);
    }
  }
}

/**
 * Function to gether necessary inputs for form creation pop-up
 */
function fetchFormCreationData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const configuration = JSON.parse(docProps.getProperty('configuration'));
    const scheduleAnalysis = analyzeScheduleData(); // Can add input here to then pass along for earliest week creation if desired.
    let apiWeek = null;
    try {
      apiWeek = fetchWeek(null, true)
    } catch (err) {
      Logger.log(`Error in fetching API week for fetchFormCreationData function: ${err.stack}`);
    }
    return {
      configuration: configuration,
      matchupData: {
        available: scheduleAnalysis.available,
        matchups: scheduleAnalysis.matchups
      },
      validitySummary: scheduleAnalysis.validitySummary,
      apiWeek: apiWeek,
      leagueData: LEAGUE_DATA,
      dayColor: dayColorsObj  || '#e0e0e0',
      dayColorBorder: dayColorsFilledObj || '#b0b0b0'
    };
  } catch (err) {
    Logger.log(`⚠️ A critical error occurred in "fetchFormCreationData": ${err.stack}`);
    // Explicitly return a safe, default object on failure.
    return { 
      configuration: {}, 
      matchupData: {
        available: false, matchups: [], },
      validitySummary: {},
      apiWeek: apiWeek,
      leagueData: LEAGUE_DATA,
      dayColor: dayColorsObj || '#e0e0e0',
      dayColorBorder: dayColorsFilledObj || '#b0b0b0'
    };
  }
}

/**
 * Fetches, filters, and analyzes schedule data to create a
 * week-by-week data quality summary.
 * @param {number} earliestWeek - The first week to include in the analysis.
 * @returns {Object} An object containing the filtered matchups and the validity summary.
 */
function analyzeScheduleData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LEAGUE); // Assumes LEAGUE is a global const like 'NFL'
    
    if (!sheet) throw new Error(`Sheet '${LEAGUE}' not found.`);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { available: false, matchups: [], validitySummary: {} };
    
    const headers = data.shift();
    const weekCol = headers.indexOf('week');
    const dateCol = headers.indexOf('date');
    const timeFetchedCol = headers.indexOf('timeFetched');
    const spreadCol = headers.indexOf('spread');
    const overUnderCol = headers.indexOf('overUnder');
    const autoFetchedCol = headers.indexOf('spreadAutoFetched');

    // Filter data to relevant weeks and convert to objects
    const matchups = data
      .filter(row => row[weekCol] >= 1) // Formerly used earliest week as input to this function and defaulted to next week after first game started
      .map(row => {
        let matchupObject = {};
        headers.forEach((header, index) => {
          let value = row[index];
          if ((index === dateCol || index === timeFetchedCol) && value instanceof Date) {
            matchupObject[header] = value.toISOString();
          } else {
            matchupObject[header] = value;
          }
        });
        return matchupObject;
      });

    if (matchups.length === 0) return { available: true, matchups: [], validitySummary: {} };

    // Group matchups by week for analysis
    const gamesByWeek = matchups.reduce((acc, game) => {
      const week = game.week;
      if (!acc[week]) acc[week] = [];
      acc[week].push(game);
      return acc;
    }, {});

    // Analyze each week's data to create the summary
    const validitySummary = {};
    for (const week in gamesByWeek) {
      const weekGames = gamesByWeek[week];
      let firstTimeFetched = weekGames[0]?.timeFetched || null;
      
      const analysis = {
        auto: weekGames.every(g => g[headers[autoFetchedCol]] === true || g[headers[autoFetchedCol]] === 1),
        timeFetched: firstTimeFetched,
        spreads: weekGames.every(g => g[headers[spreadCol]] !== ''),
        overUnders: weekGames.every(g => g[headers[overUnderCol]] !== '')
      };

      // Check for inconsistent fetch times
      const isTimeConsistent = weekGames.every(g => g.timeFetched === firstTimeFetched);
      if (!isTimeConsistent) {
        analysis.timeFetched = 'ERROR';
      }
      
      validitySummary[week] = analysis;
    }

    return {
      available: true,
      matchups: matchups,
      validitySummary: validitySummary
    };
  } catch (e) {
    Logger.log(`⚠️ Error in analyzeScheduleData: ${err.stack}`);
    return { available: false, matchups: [], validitySummary: {} };
  }
}

/**
 * Updates the 'Schedule' sheet with user-provided override data.
 * @param {number} week The week being updated.
 * @param {Object} customData The object of user edits from the client.
 */
function updateScheduleData(week, customData) {
  if (!customData || Object.keys(customData).length === 0) {
    return; // Nothing to update
  }
  
  Logger.log(`✏️ Applying ${Object.keys(customData).length} user overrides for Week ${week}.`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LEAGUE); // e.g., 'NFL'
  if (!sheet) throw new Error(`Sheet '${LEAGUE}' not found.`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const weekCol = headers.indexOf('week');
  const awayCol = headers.indexOf('awayTeam');
  const homeCol = headers.indexOf('homeTeam');
  const spreadCol = headers.indexOf('spread');
  const overUnderCol = headers.indexOf('overUnder');
  const timeFetchedCol = headers.indexOf('timeFetched');

  let changesMade = 0;
  
  // Loop through the data array (skipping headers)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[weekCol] == week) {
      const gameId = `${row[awayCol]}@${row[homeCol]}`;
      if (customData[gameId]) {
        const edits = customData[gameId];
        if (edits.spread !== undefined) {
          data[i][spreadCol] = edits.spread;
          changesMade++;
        }
        if (edits.overUnder !== undefined) {
          data[i][overUnderCol] = edits.overUnder;
          changesMade++;
        }
        // Also update the timeFetched and auto-fetch status to reflect manual edit
        data[i][timeFetchedCol] = new Date();
        // Assuming you have a 'spreadAutoFetched' column
        const autoFetchedCol = headers.indexOf('spreadAutoFetched');
        if (autoFetchedCol > -1) {
          data[i][autoFetchedCol] = 0; // It's now a manual override
        }
      }
    }
  }

  // If changed, write the entire data range back to the sheet.
  if (changesMade > 0) {
    sheet.getDataRange().setValues(data);
    Logger.log(`✅ Successfully updated the Schedule sheet.`);
  }
}

/**
 * Receives the "game plan" from the UI, performs pre-flight checks,
 * confirms with the user if necessary, and then executes the form creation.
 *
 * @param {Object} gamePlan The detailed object of user intent from the client.
 */
function createNewFormForWeek(gamePlan) {
  const ui = fetchUi();
  const ss = fetchSpreadsheet();
  const week = gamePlan.week;
  
  // Preliminary Checks
  const docProps = PropertiesService.getDocumentProperties();
  const forms = JSON.parse(docProps.getProperty(`forms`)) || {};
  const config = JSON.parse(docProps.getProperty('configuration'));

  if (!gamePlan.hasOwnProperty('edits') || gamePlan?.edits) {
    try {
      updateScheduleData(ss, gamePlan);
      ss.toast(`Placed your modified spread data if present for week ${week}.`,`📋 UPDATED SPREAD DATA`)
    } catch (err) {
      Logger.log(`Issue passing the new spread data to the ${LEAGUE} sheet, moving on (these spreads are saved with the 'forms' variable).`)
    }
  } else if (config.pickemsAts || config.survivorAts || config.eliminatorAts) {
    ss.toast(`Using the existing spread data from the ${LEAGUE} sheet for week ${week}.`,`✅ USING EXISTING SPREADS`);
  }

  
  let warnings = [];
  if (forms[week]) {
    if (forms[week].formId) {
      warnings.push("An existing form for this week will be sent to your trash folder.\n\nAny associated responses in the database file will be archived.");
    }
  }

  // User Confirmation (if necessary)
  if (warnings.length > 0) {
    const message = warnings.join("\n") + "\n\nAre you sure you want to proceed?";
    const response = ui.alert(`⚠️ WARNING!`,message,ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      throw new Error("Form creation canceled by user.");
    }
  }

  try {
    // Prepare for Creation
    if (forms[week] && forms[week].formId) {
      try {
        DriveApp.getFileById(forms[week].formId).setTrashed(true);
      } catch (e) {
        Logger.log(`Could not trash old form with ID ${forms[week].formId}. It may have been deleted already. Continuing.`);
      }
    }
    // Clear out old data from properties
    delete forms[week];
  } catch (err) {
    Logger.log(`Failed to create form for week ${week}:`, err.stack);
    throw new Error(`Failed to create form: ${err.stack}`);
  }
  // Execute the "Worker"
  const newFormDetails = buildFormFromGamePlan(gamePlan);
  try {
    if (newFormDetails.formId) {
      // Record the current state of these properties for data fetching integrity
      gamePlan.pickemsInclude = config.pickemsInclude;
      gamePlan.tiebreakerInclude = config.tiebreakerInclude;
      // Pin the tiebreaker game into the saved plan so pick import and outcome
      // placement both act on the same game the form actually asked about.
      if (config.tiebreakerInclude) {
        const tiebreakerGame = getTiebreakerGame(gamePlan.games);
        if (tiebreakerGame) {
          gamePlan.tiebreakerGame = {
            awayTeam: tiebreakerGame.awayTeam,
            homeTeam: tiebreakerGame.homeTeam,
            shortName: `${tiebreakerGame.awayTeam} @ ${tiebreakerGame.homeTeam}`,
            dayName: tiebreakerGame.dayName,
            hour: tiebreakerGame.hour,
            minute: tiebreakerGame.minute
          };
        }
      }
      gamePlan.survivorInclude = newFormDetails.survivorInclude;
      gamePlan.eliminatorInclude = newFormDetails.eliminatorInclude;
      gamePlan.pickemsAts = config.pickemsAts;
      gamePlan.survivorAts = config.survivorAts;
      gamePlan.eliminatorAts = config.eliminatorAts;

      // Save the new state
      const newFormsData = {
        formId: newFormDetails.formId,
        editUrl: newFormDetails.editUrl,
        publishedUrl: newFormDetails.publishedUrl,
        active: true,
        respondents: [],
        nonRespondents: newFormDetails.eligibleMembers,
        gamePlan: gamePlan
      };
      forms[week] = newFormsData;
      forms[week].imported = false;

      // Storing properties
      saveProperties('forms',forms);
      
      // Setting up synce for form
      try {
        setFormSubmitTrigger(newFormDetails.formId, true)
      } catch (err) {
        Logger.log(`⚠️ Could not set up trigger for new week ${week} form: ${err.stack}`);
      }

      // Schedule the automatic nudge to whoever has not submitted
      try {
        scheduleWeeklyReminder(week, gamePlan);
      } catch (err) {
        Logger.log(`⚠️ Could not schedule the week ${week} reminder: ${err.stack}`);
      }

      // Setting up onEdit trigger if includes pool questions
      try {
        if (newFormDetails.survivorInclude || newFormDetails.eliminatorInclude) {
          createOnEditTrigger();
          Logger.log(`✅ Created onEdit trigger to continue updates to contests.`);
        }
      } catch (err) {
        Logger.log(`⚠️ Error setting onEdit trigger for pools. Try again manually or get help.`)
        ss.toast(`Issue setting an onEdit trigger for the contest pools, please try manually or reach out for help.`,`⚠️ ON EDIT TRIGGER ERROR`)
      }
      
      showFormActionsDialog(newFormsData, week);

      return { success: true, message: `✅ Successfully created form for week ${week}.` };
    } else {
      Logger.log(`⚠️ Encountered an error with the form creation during the 'buildFormFromGamePlan' function`);
    }
  } catch (err) {
    Logger.log(`⚠️ Despite making a form, an error was encountered during wrap-up. | ERROR: ${err.stack}`);
  }
}

/**
 * Fetches, validates, or creates the single backend "database" spreadsheet.
 * This is the reciprocal to your formTemplate function.
 *
 * @returns {Spreadsheet} The active, valid Spreadsheet object for the database.
 */
function getDatabaseSheet() {
  const docProps = PropertiesService.getDocumentProperties();
  const dbId = docProps.getProperty('databaseId');

  if (dbId) {
    try {
      // Try to open the spreadsheet using the stored ID.
      const file = DriveApp.getFileById(dbId);
      if (file.getName() && !file.isTrashed()) {
        return SpreadsheetApp.openById(dbId);
      } else {
        docProps.deleteProperty('databaseId');
      }
    } catch (err) {
      // This catch block runs if the file was deleted or permissions changed.
      Logger.log(`Could not open database sheet with ID "${dbId}". It may have been deleted. A new one will be created.`);
      docProps.deleteProperty('databaseId'); // Clear the invalid ID
    }
  }

  // If we reach here, either there was no ID or the old one was invalid.
  const ui = SpreadsheetApp.getUi();
  ui.alert('Backend Database Not Found', 'A private spreadsheet for storing form responses could not be found. A new one will be created now.', ui.ButtonSet.OK);
  
  const config = JSON.parse(docProps.getProperty('configuration'));
  const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);

  const dbName = `DO NOT DELETE - ${fetchYear()} ${config.groupName || LEAGUE + ' Picks Pool - Form Responses'}`;
  const newDb = SpreadsheetApp.create(dbName);
  const newDbId = newDb.getId();
  
  // Move to the forms folder as well
  DriveApp.getFileById(newDbId).moveTo(formsFolder);
  
  // Add a helpful note for the user in the new sheet.
  newDb.getSheets()[0].getRange('A1').setValue(`This sheet is the private backend for the ${config.groupName} pool. Please do not delete, rename, or share this file.`);

  docProps.setProperty('databaseId', newDbId);
  Logger.log(`Created new database sheet with ID: ${newDbId}`);
  
  return newDb;
}

/**
 * Fetches, validates, or creates the template form.
 * Now prompts the user with a YES/NO choice if the template needs to be created.
 *
 * @returns {Form|null,boolean} The Form object, or null if the user cancels and a boolean if the form ID stored didn't work upon onset
 */
function getTemplateForm() {
  Logger.log(`▶️ Beginning process to get template form...`);
  const docProps = PropertiesService.getDocumentProperties();
  const templateId = docProps.getProperty('templateId');
  if (templateId) {
    Logger.log(`🔎 Found template id: ${templateId}`);
    try {
      const file = DriveApp.getFileById(templateId);
      if (file && !file.isTrashed()) {
        Logger.log(`🔎 Found an existing template file, attempting to open...`);
        return FormApp.openById(templateId);
      }
    } catch (err) {
      Logger.log(`⚠️ Could not open template form with ID "${templateId}". ${err.stack}`);
      docProps.deleteProperty('templateId');
    }
  }
  const templateName = 'PICKS TEMPLATE - Customize Header and Color';
  const newTemplate = FormApp.create(templateName);
  const newTemplateId = newTemplate.getId();
  
  // Ensure the template is moved into the correct folder.
  const config = fetchProperties('configuration');
  const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);
  DriveApp.getFileById(newTemplateId).moveTo(formsFolder);

  docProps.setProperty('templateId', newTemplateId);
  Logger.log(`🌟 Created new template form with ID: ${newTemplateId}`);

  return newTemplate;
}

/**
 * Fetches, validates, or creates the folder for storing weekly forms.
 *
 * @param {string} groupName The name of the pool, passed in to avoid extra fetches.
 * @returns {Folder} The active, valid Folder object.
 */
function getFormsFolder(groupName) {
  const docProps = PropertiesService.getDocumentProperties();
  const folderId = docProps.getProperty('folderId');
  
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      // A simple check to ensure it's a valid folder.

      if (folder.getName() && !folder.isTrashed()) {
        return folder;
      }
    } catch (err) {
      Logger.log(`Could not open forms folder with ID "${folderId}". It may have been deleted. A new one will be created.${err.stack}`);
      docProps.deleteProperty('folderId');
    }
  }

  const folderName = `${fetchYear()} ${groupName ? groupName : LEAGUE + ' Picks Pool'} Forms`;
  const newFolder = DriveApp.createFolder(folderName);
  docProps.setProperty('folderId', newFolder.getId());
  Logger.log(`Created new forms folder: "${folderName}"`);
  
  return newFolder;
}

function checkCONFIGDATA() {
  const formsData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('configuration'));
  Logger.log(Object.keys(formsData).length);
  Logger.log(JSON.stringify(formsData.survivorActive));
  Logger.log(JSON.stringify(formsData.eliminatorActive));
  Logger.log(JSON.stringify(formsData));
}
/**
 * Takes a final, validated "game plan" and builds a Google Form.
 * This function is now modular and uses the ID-based members object.
 *
 * @param {Object} gamePlan The detailed plan for the form.
 * @returns {Object} An object with the new form's ID and URLs.
 */
function buildFormFromGamePlan(gamePlan) {
  const ss = fetchSpreadsheet();
  try {
    // --- Setup and Variable Initialization ---
    const docProps = PropertiesService.getDocumentProperties();
    const config = JSON.parse(docProps.getProperty('configuration'));
    const formsData = JSON.parse(docProps.getProperty('forms'));
    let memberData = JSON.parse(docProps.getProperty('members'));
    const week = parseInt(gamePlan.week, 10);
    const formName = gamePlan.formName;

      // 1. Get the validated helper objects.
    const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);
    const templateForm = getTemplateForm();
    let databaseSheet = getDatabaseSheet();
    
    // 2. Create the new form by copying the template.
    const newFormFile = DriveApp.getFileById(templateForm.getId()).makeCopy(formName, formsFolder);
    const form = FormApp.openById(newFormFile.getId()).setPublished(true);

    const urlFormEdit = form.shortenFormUrl(form.getEditUrl());
    const urlFormPub = form.shortenFormUrl(form.getPublishedUrl());
    ss.toast(`Created form and generated links`,`✅ EMPTY FORM CREATED`);
    Logger.log(`✅ Created empty form and generated links`);
    
    // FORM TO DATABASE SHEET LINKING
    // Process to link spreadsheet created to backend database:
    // 1. Pull sheets within database file
    // 2. Create destination sheet
    // 3. Review, with periodic waits, for a new sheet to appear within sheets listing
    // 4. ID the sheet that is new
    // 5. Archive any sheets that share the WK${week} name
    // 6. Rename new sheet to WK${name}
    formDatabaseLinking(week,form,databaseSheet,ss);

    // ESTABLISH ALL STATES
    const pickems = config.pickemsInclude;

    const survivorStart = parseInt(config.survivorStartWeek,10) == week;
    let survivor = config.survivorInclude && week >= parseInt(config.survivorStartWeek,10);
    
    const eliminatorStart = parseInt(config.eliminatorStartWeek,10) == week;
    let eliminator = config.eliminatorInclude && week >= parseInt(config.eliminatorStartWeek,10);
    let hasMembers = false;
    try {
      hasMembers = memberData.memberOrder && memberData.memberOrder.length > 0;
    } catch (err) {
      Logger.log(`🚫 No member data found, this will be a one page form.`);
    }
    const firstWeek = formsData ? (Object.keys(formsData).length > 0 ? false : true) : true;
    if (firstWeek) {
      ss.toast(`No other forms detected, assuming this week (${week}) is the start of your group`,`🚀 START WEEK DETECTED`);
      Logger.log(`🚀 No other forms detected, assuming this week (${week}) is the start of your group`);
    }
    // FORM BUILDING ROUTER

    /// 1. NAME QUESTION
    let nameQuestion, welcomeHeader;
    // In the event of no members, we shift away from name choices and simply create a text box for name entry to start the pool up
    if (hasMembers) {
      nameQuestion = form.addListItem()
        .setTitle('Select Your Name')
        .setRequired(true);
      if (gamePlan.membershipLocked) {
        nameQuestion.setHelpText('✏️ New users will be prompted to enter a name on the next page');
      }
    } else {
      nameQuestion = form.addTextItem()
        .setTitle('Enter Your Name')
        .setHelpText('Please enter your name as it will appear in the pool.')
        .setRequired(true)
        .setValidation(nameValidation); // Validation is with constants at beginning of document
    }
    if (firstWeek || week == 1) {
      const text = week >= 19 ? `👋 Welcome to the ${config.year} ${LEAGUE} playoffs!` : `👋 Welcome to the ${config.year} ${LEAGUE} season!`
      welcomeHeader = form.addSectionHeaderItem()  
        .setTitle(text);
      if (config.welcomeLetter) {
        welcomeHeader.setHelpText(config.welcomeLetter);
      }
    }

    /// 2. PICK 'EMS
    if (pickems) {
      ss.toast(`Creating pick 'ems questions for week ${week}`,`🏈 PICK 'EMS`);
      // 1. Add a Section Header to act as the title for this part of the form.
      form.addSectionHeaderItem().setTitle("🏈 Weekly Pick 'Em Selections");
      // 2. If ATS is enabled for Pick'em, add a very clear instructional message.
      if (config.pickemsAts) {
        form.addSectionHeaderItem()
          .setTitle('🔢 Instructions: Pick Against the Spread (ATS)')
          .setHelpText('For each game, select the team you believe will win WITH the point spread. The point spread is listed in the help text of each question.');
      }
      buildPickemQuestions(ss, form, gamePlan, config);
      ss.toast(`Created pick 'ems questions for week ${week}`,`✅ PICK 'EMS DONE`);
    } else {
      Logger.log(`❌ No pick 'ems pool active, moving on to survivor/eliminator`)
    }

    // --- Build Survivor/Eliminator Pages ---
    let singlePageForm = false, submitPage, nameChoices = [], eligibleMembers = [];
    if (hasMembers) {
      submitPage = form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
    } else {
      singlePageForm = true;
      Logger.log(`📃 Single Page form being deployed -- no members to allow for selection.`);
    }
    const contests = (survivor && eliminator) ? 2 : (survivor || eliminator ? 1 : 0);
    if (contests > 0) {
      for (let a = config.survivorStart; a < week; a++) {
        evalSurvElimStatus(a);
      }
      let survivorMembers = 0, eliminatorMembers = 0;
      let allSurvivorTeamsForWeek = buildTeamList(gamePlan, config, config.survivorAts);
      let allEliminatorTeamsForWeek = buildTeamList(gamePlan, config, config.eliminatorAts);
      if (((survivor && week === parseInt(config.survivorStartWeek, 10)) && (eliminator && week === parseInt(config.eliminatorStartWeek, 10))) ||
          (survivor && week === parseInt(config.survivorStartWeek, 10) && !eliminator) ||
          (eliminator && week === parseInt(config.eliminatorStartWeek, 10) && !survivor)) {
        // Create ONE common page for all existing users for the first week.
        form.addSectionHeaderItem().setTitle(`🏆 Contest Pick${survivor && eliminator ? 's' : ''}`);
        if (survivor) {
          addContestQuestion(form, 'survivor', {}, config.survivorAts, config.survivorStartWeek, allSurvivorTeamsForWeek, survivorStart); // Pass empty member object
        }
        if (eliminator) {
          addContestQuestion(form, 'eliminator', {}, config.eliminatorAts, config.eliminatorStartWeek, allEliminatorTeamsForWeek, eliminatorStart); // Pass empty member object
        }
        Logger.log(`1️⃣ First week of contest(s), creating a generic drop-down for included games`);
        ss.toast(`First week of one of your contests where the other is inactive/absent, created generic question(s).`,`1️⃣ FIRST WEEK FOR CONTEST(S)`);
        if (hasMembers) {
          memberData.memberOrder.forEach(memberId => {
            const member = memberData.members[memberId];
            if (member) {
              Logger.log(`➕ Adding first week name choice option for member: ${member.name}.`)
              nameChoices.push(nameQuestion.createChoice(member.name, submitPage));
            } else {
              Logger.log(`❗ Unable to add first week name choice option for member: ${member.name}.`)
              Logger.log(`❔ Member Data: \n ${JSON.stringify(member)}`)
            }
          });
          ss.toast(`Generic contest page(s) created and linked members to submit page.`,`🔀 MEMBERS ROUTED`);
          Logger.log(`🔀 Linked all members from generic page to the submit page... adding new user page if needed.`);
        }
      } else if (hasMembers) {
        const sLivesIndex = Math.max(0,week - 2); // e.g., for Week 2, check index 0.
        const eLivesIndex = Math.max(0,week - 2);
        const sLS = config.survivorLives;
        const eLS = config.eliminatorLives;
        ss.toast(`Creating questions for members who are in both survivor and eliminator for week ${week}`,`👑&💀 SURVIVOR AND ELIMINATOR`);
        Logger.log(`👑&💀 Creating possible destinations for instances where both Survivor and Eliminator are both active`)
        memberData.memberOrder.forEach(memberId => {
          const member = memberData.members[memberId];
          if (member) {
            Logger.log(`🔹${member.name} Data${survivor ? '\nSurvivor Lives: ' + member.sL : ''}${eliminator ? '\nEliminator Lives: ' + member.eL : ''}`);
            let contestIcon, contestText, helpText, survivorHelp, eliminatorHelp, both = false;            
            const survivorStatus = survivor && isMemberEligible(member, 'survivor', week, config);
            const eliminatorStatus = eliminator && isMemberEligible(member, 'eliminator', week, config);
        
            if (survivorStatus || eliminatorStatus) {
              if (survivorStatus && eliminatorStatus) {
                contestIcon = `👑&💀`;
                contestText = `${member.name} is active in SURVIVOR & ELIMINATOR`;
                survivorMembers++;
                eliminatorMembers++;
                both = true;
              } else if (survivorStatus) {
                contestIcon = `👑`;
                contestText = `${member.name} is active in SURVIVOR ${contests > 1 ? 'ONLY' : ''}`;
                survivorMembers++;
              } else if (eliminatorStatus) {
                contestIcon = `💀`;
                contestText = `${member.name} is active in ELIMINATOR ${contests > 1 ? 'ONLY' : ''}`;
                eliminatorMembers++;
              }
              ss.toast(contestText,`${contestIcon} CREATING CONTEST QUESTION`);
              Logger.log(`${contestIcon} Creating Customized Page: ${contestText}`);
              
              if (survivor) survivorHelp = sLS == 1 ? `One Survivor Life: ${createLivesString(member.sL[sLivesIndex],sLS)}` : `Survivor Lives: ${createLivesString(member.sL[sLivesIndex],sLS)} (${member.sL[sLivesIndex] < sLS ? member.sL[sLivesIndex] + ' remaining' : 'all remaining'})`;
              if (eliminator) eliminatorHelp = eLS == 1 ? `One Eliminator Life: ${createLivesString(member.eL[eLivesIndex],eLS)}` : `Eliminator Lives: ${createLivesString(member.eL[eLivesIndex],eLS)} (${member.eL[eLivesIndex] < eLS ? member.eL[eLivesIndex] + ' remaining' : 'all remaining'})`;
              helpText = contests > 1 ? `${survivorHelp}  |  ${eliminatorHelp}` : (survivor ? survivorHelp : eliminatorHelp);
              
              const title = `${member.name}'s ${both ? 'Survivor & Eliminator' : (survivorStatus ? 'Survivor' : 'Eliminator')} Pick${both ? 's' : ''}`;
              
              const contestPage = form.addPageBreakItem().setTitle(title).setHelpText(helpText).setGoToPage(FormApp.PageNavigationType.SUBMIT);
              
              // Add both questions to this single page
              if (survivorStatus) addContestQuestion(form, 'survivor', member, config.survivorAts, config.survivorStartWeek, allSurvivorTeamsForWeek, survivorStart);
              if (eliminatorStatus) addContestQuestion(form, 'eliminator', member, config.eliminatorAts, config.eliminatorStartWeek, allEliminatorTeamsForWeek, eliminatorStart);
    
              nameChoices.push(nameQuestion.createChoice(member.name, contestPage));
              
              Logger.log(`✅ Question${both ? 's' : ''} for ${member.name} created.`);
            } else {
              Logger.log(`❌ ${member.name} Is out or ineligible for ${contests > 1 ? 'both contests' : 'the contest'}.`)
              nameChoices.push(nameQuestion.createChoice(member.name, submitPage));
            }
          } else {
            Logger.log(`⚠️ Invalid member! (ID: ${memberId})`)
          }
        });
      } else {
        ui.alert(`⚠️ Start week past with no members!`,`You likely are seeing this because you've set a start date for a survivor or eliminator pool that is prior to this week's form and have no active members who would be eligible. Please adjust your start week in the configuration tool (⚙️) to this current week (${week}) and try again.`, ui.ButtonSet.OK);
        return;
      }
      ss.toast(`Created all contest questions.${survivor ? '\nSurvivor Members: ' + survivorMembers : ''}${eliminator ? '\nEliminator Members: ' +  eliminatorMembers : ''}`,`🔀 MEMBERS ROUTED`);
      Logger.log(`🔀 Linked all members to their respective pages for navigation... adding new user page if needed.${survivor ? '\nSurvivor Members: ' + survivorMembers : ''}${eliminator ? '\nEliminator Members: ' +  eliminatorMembers : ''}`);
    } else if (hasMembers) { // Event where contests didn't exist or are over and need to route all members to a submit page instead of contest pages
      Logger.log(`❌ No survivor or eliminator contest active`);
      memberData.memberOrder.forEach(memberId => {
        const member = memberData.members[memberId];
        if (member) {
          Logger.log(`➕ Adding name choice option for member: ${member.name}.`)
          nameChoices.push(nameQuestion.createChoice(member.name, submitPage));
        } else {
          Logger.log(`❗ Unable to add name choice option for member: ${member.name}.`)
          Logger.log(`❔ Member Data: \n ${JSON.stringify(member)}`)
        }
      });
      ss.toast(`Completed form questions and routed all existing members to submit page.`,`🔀 MEMBERS ROUTED`);
      Logger.log(`🔀 Linked all members to submit page... adding new user page if needed.`);
    }

    // Add 'New User' option if applicable
    if (!config.membershipLocked) {
      const text = 'Membership is unlocked--creating a new user question';
      Logger.log(`🔓 ${text}`);
      ss.toast(text,`🔓 MEMBERSHIP UNLOCKED`)
      if (!singlePageForm) {
        const newUserPage = buildNewUserPage(ss, form, config, gamePlan, survivor, eliminator, survivorStart, eliminatorStart);
        nameChoices.unshift(nameQuestion.createChoice(('✏️ NEW USER'), newUserPage));
      }
    } else {
      const text = 'Membership is locked--no new user question added';
      Logger.log(`🔓 ${text}`);
      ss.toast(text,`🔒 MEMBERSHIP LOCKED`)
    }
    
    if (!singlePageForm) { // If single page form there is only a text entry field and a submit button on bottom.
      if (nameChoices.length === 0) { // Check for failure to make name choices - probably unnecessary now
        nameChoices.push(nameQuestion.createChoice("No members eligible", submitPage));
        form.setDescription(`⚠️ Warning: No members are currently eligible to make picks. Please check Member Management or unlock membership.`);
      } else {
        Logger.log(`📝 Setting Name Choices`);
        nameQuestion.setChoices(nameChoices);
        ss.toast('Set all choices for name question drop-down','📝 NAME CHOICES SET');
      }
    }

    // --- Final Touches ---
    Logger.log(`↩️ Returning information to form creation controller...`);
    ss.toast(`Returning information to form creation controller...`,`↩️ REROUTING DATA`);
    
    const formId = form.getId();
    if (config.kickoffLock) setOneTimeFormLockTrigger(formId, gamePlan);
    
    return {
      formId: formId,
      editUrl: urlFormEdit,
      publishedUrl: urlFormPub,
      eligibleMembers: eligibleMembers,
      survivorInclude: survivor,
      eliminatorInclude: eliminator
    };
  } catch (err) {
    ss.toast(`Encountered an issue during the creation of the form. Check logs for details: ${err.stack}`,`❗ ERROR CREATING FORM`)
    Logger.log(`❗ Encountered an issue during the creation of the form: ${err.stack}`)
    Logger.log(`❌ Deleting form if it was created somewhere during the process...`);
    try {
      DriveApp.getFileById(form.getId()).setTrashed(true);
      Logger.log(`🗑 Successfully trashed form.`)
    } catch (err) {
      Logger.log(`❌ Unable to trash new form or it didn't exist yet.`)
    }
    return err;
  }
}

/**
 * Eligibilty checking for members
 */
function isMemberEligible(member, contestType, week, config) {
  if (!member || !member.active) return false;
  const livesKey = contestType === 'survivor' ? 'sL' : 'eL';
  const livesData = member[livesKey];
  const startWeek = parseInt(config[`${contestType}StartWeek`], 10) || 1;
  // For the official start week, all active members are eligible.
  if (week === startWeek) {
    return true;
  }
  // For subsequent weeks, check their life status from the previous week.
  if (Array.isArray(livesData)) {
    // New "lives array" format
    return livesData[week - 2] > 0;
  } else {
    // Old number format (fallback)
    return livesData > 0;
  }
}

/**
 * Form has a back-end database that stores its responses immediately for faster fetching.
 * This document ensures that the pool members don't see the responses until they've been imported
 * Newly created forms will generate their own new sheet within the document when linked
 * The new sheet needs to be renamed to ensure the data goes where we can fetch it
 */
function formDatabaseLinking(week,form,databaseSheet,ss) {
  week = week || fetchWeek();
  ss = ss || fetchSpreadsheet();
  databaseSheet = databaseSheet || getDatabaseSheet();
  form = form || FormApp.openById(JSON.parse(PropertiesService.getDocumentProperties().getProperty('forms'))[week].formId);
  if (form) {
    try {
      const databaseSheetId = databaseSheet.getId();
      
      // Get the initial count of sheets before linking
      const initialSheets = databaseSheet.getSheets();
      const initialSheetNames = initialSheets.map(sheet => sheet.getName());
      // Set the form's destination. This creates a new sheet in the spreadsheet.
      form.setDestination(FormApp.DestinationType.SPREADSHEET, databaseSheetId);
      let waitPeriod = 1;
      let updatedSheets = [];
      let updatedSheetNames = [];
      let foundNewSheet = false;
      while (waitPeriod <= 40 && !foundNewSheet) {
        // Wait a moment for Google to create the new sheet...
        Utilities.sleep(500);
        SpreadsheetApp.flush();
        
        // Re-fetch the spreadsheet to get the current state
        databaseSheet = SpreadsheetApp.openById(databaseSheetId);
        
        // Get the updated list of sheets
        updatedSheets = databaseSheet.getSheets();
        updatedSheetNames = updatedSheets.map(sheet => sheet.getName());
        
        Logger.log(`Attempt ${waitPeriod}: Found ${updatedSheets.length} sheets: [${updatedSheetNames.join(', ')}]`);
        
        // Check if we have a new sheet
        if (updatedSheets.length > initialSheets.length) {
          Logger.log(`Success! Found new sheet after ${waitPeriod} attempts.`);
          foundNewSheet = true;
          break; // Exit the loop immediately
        }
        waitPeriod++;
      }
      if (!foundNewSheet) {
        throw new Error(`Timed out waiting for new response sheet to be created after ${waitPeriod - 1} attempts.`);
      }
    
      let newResponseSheet = null;

      // Find the sheet that wasn't there before
      for (const sheet of updatedSheets) {
        if (!initialSheetNames.includes(sheet.getName())) {
          newResponseSheet = sheet;
          break; // Exit as soon as we find it
        }
      }
        
      // Fallback - look for sheets with "Form Responses" pattern
      if (!newResponseSheet) {
        Logger.log(`⏮ Fallback: Looking for 'Form Responses' pattern...`);
        for (const sheet of updatedSheets) {
          if (sheet.getName().includes('Form Responses')) {
            // Check if this one is new by comparing against initial list
            if (!initialSheetNames.includes(sheet.getName())) {
              newResponseSheet = sheet;
              break;
            }
          }
        }
      }
      if (!newResponseSheet) {
        // Debug: Show exactly what we found
        Logger.log(`DEBUG - Initial sheets: ${JSON.stringify(initialSheetNames)}`);
        Logger.log(`DEBUG - Updated sheets: ${JSON.stringify(updatedSheetNames)}`);
        Logger.log(`DEBUG - Difference: " ${JSON.stringify(updatedSheetNames.filter(name => !initialSheetNames.includes(name)))}`);
        throw new Error(`⚠️ Could not identify the newly created response sheet.`);
      }
      Logger.log(`Found new response sheet: "${newResponseSheet.getName()}"`);

      // Now let's rename and organize it
      const newSheetName = `WK${week}`;
      
      // Check if a sheet with the desired name already exists
      const existingSheet = databaseSheet.getSheetByName(newSheetName);
      
      if (existingSheet) {
        Logger.log(`An existing sheet named '${newSheetName}' was found. Archiving it now.`);
        ss.toast(`Found a former WK${week} sheet in the database, archiving now`,`💾 OLD RESPONSES ARCHIVED`);
        let archiveIndex = 1;
        let archiveName = `WK${week}_ARCHIVE${archiveIndex}`;
        
        // Loop to find a unique archive name that doesn't already exist
        while (databaseSheet.getSheetByName(archiveName) !== null) {
          archiveIndex++;
          archiveName = `WK${week}_ARCHIVE${archiveIndex}`;
        }
        
        // Rename and hide the old sheet
        existingSheet.setName(archiveName);
        existingSheet.hideSheet();
        
        Logger.log(`Successfully archived old sheet as '${archiveName}'.`);
      }
      
      // Now safely rename the newly linked sheet
      newResponseSheet.setName(newSheetName);
      newResponseSheet.activate(); // Brings the new tab to the front
      Logger.log(`🔗 Successfully linked form and renamed response sheet to '${newSheetName}'.`);
      ss.toast(`Backend database new response sheet found and renamed to WK${week}`,`🔗 DATABASE LINKED`);
      return { status: 'success'};
    } catch (err) {
      // If linking fails, delete the form to avoid orphans
      DriveApp.getFileById(form.getId()).setTrashed(true);
      Logger.log(`Failed to link form to spreadsheet and manage tab. Form deleted. Error: ${err.stack}`);
      throw new Error("Could not link form to the backend database. Please check permissions.");
      // return { status: 'failed' }
    }
  } else {
    Logger.log(`⭕ No form passed to database linking function, try again later`);
  }
}


/**
 * Adds a single, correctly filtered contest question to the form.
 */
function addContestQuestion(form, contestType, member, isAts, startWeek, allTeamsForWeek, isStart) {
  let availableTeams = [];
  if (isStart) {
    availableTeams = [...allTeamsForWeek];
  } else {
    const startIndex = (parseInt(startWeek, 10) || 1) - 1;
    const picksKey = contestType === 'survivor' ? 'sP' : 'eP';
    const allMemberPicks = member[picksKey] || [];
    const relevantPicks = allMemberPicks.slice(startIndex).filter(team => {
      const teamAbbr = team.split(' ')[0];
      return teamAbbr;
    });
    
    availableTeams = allTeamsForWeek.filter(team => {
      const teamAbbr = team.split(' ')[0];
      return !relevantPicks.includes(teamAbbr);
    });
  }
  if (availableTeams.length === 0) {
    Logger.log(`Member has already selected all available teams for this week of the formerly available ${allTeamsForWeek.length}. Group owner should adjust form`);
    SpreadsheetApp.getUi().alert(`One of your members has no avaialble options for your current form selections of games to include. Please re-run the form builder and select more games.`);
    availableTeams = ['NO OPTIONS'];
  }

  let title = (contestType === 'survivor' ? '👑 ' : '💀 ') + capitalize(contestType) + (contestType === 'eliminator' ? ' Loser' + (isAts ? ' ATS' : '') + ' Pick' : ' Winner' + (isAts ? ' ATS' : '') + ' Pick');
  if (member.name) title += ` (${member.name})`;
  const helpText = `Select which team you believe will ${isAts ? (contestType === 'survivor' ? 'WIN' : 'LOSE') + ' when factoring in the given spread' : (contestType === 'survivor' ? 'WIN' : 'LOSE') + ' this week'}.`;
  form.addListItem()
    .setTitle(title)
    .setHelpText(helpText)
    .setChoiceValues(availableTeams)
    .setRequired(true);
}

/**
 * Builds all Pick'em related questions on the form.
 */
function buildPickemQuestions(ss, form, gamePlan, config) {
  // Chosen once, up front. This previously fell out of the game loop below, which
  // made the tiebreaker whichever game happened to be last in the game plan.
  const tiebreakerGame = config.tiebreakerInclude ? getTiebreakerGame(gamePlan.games) : null;
  if (config.tiebreakerInclude && !tiebreakerGame) {
    Logger.log(`⚠️ Tiebreakers are enabled but no tiebreaker game could be chosen for week ${gamePlan.week}`);
  }
  Logger.log(`🏈 Building Pick'em questions...`);
  gamePlan.games.forEach(game => {
    let item = form.addMultipleChoiceItem();
    const evening = game.hour >= 17;
    const mnf = evening && game.dayName === "Monday";
    let title = `${game.awayTeamLocation} ${game.awayTeamName} at ${game.homeTeamLocation} ${game.homeTeamName}${game.divisional == 1 && game.division ? ' ('+game.division+' Divisional Game)':''}`;
    let helpText = `${mnf ? 'Monday Night Football' : game.dayName} at ${formatTime(game.hour, game.minute)}`;
    if (config.pickemsAts && game.spread) helpText += `  | ↔️ Spread: ${game.spread}`;
    if (game.bonus > 1) title += ` (${game.bonus == 3 ? '3️⃣' : '2️⃣'}x Bonus)`;
    item.setTitle(title)
      .setHelpText(helpText)
      .setChoices([
        item.createChoice(`${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.awayTeam].mascot: ''} ${game.awayTeam}`), // + LEAGUE_DATA[game.awayTeam].colors_emoji 
        item.createChoice(`${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.homeTeam].mascot: ''} ${game.homeTeam}`)]) // + LEAGUE_DATA[game.homeTeam].colors_emoji 
      .showOtherOption(false)
      .setRequired(true);
    ss.toast(`Added pick 'ems question of ${game.awayTeam} @ ${game.homeTeam}`,`${LEAGUE_DATA[game.awayTeam].mascot}@${LEAGUE_DATA[game.homeTeam].mascot}`);
    Logger.log(`🏈 Pick 'Ems: ${LEAGUE_DATA[game.awayTeam].mascot}@${LEAGUE_DATA[game.homeTeam].mascot} created`);
  });
  if (config.tiebreakerInclude && tiebreakerGame) { // Excludes tiebreaker questions if tiebreakers are disabled
    const tiebreakerMatchup = `${tiebreakerGame.awayTeamLocation} ${tiebreakerGame.awayTeamName} at ${tiebreakerGame.homeTeamLocation} ${tiebreakerGame.homeTeamName}`;
    const tiebreakerOverUnder = tiebreakerGame.overUnder;
    const mnf = tiebreakerGame.dayName === 'Monday';

    const combinedValidation = FormApp.createTextValidation()
      .setHelpText('Input must be a whole number between 0 and 120')
      .requireWholeNumber()
      .requireNumberBetween(0,120)
      .build();
    const winnerValidation = FormApp.createTextValidation()
      .setHelpText('Input must be a whole number between 0 and 100')
      .requireWholeNumber()
      .requireNumberBetween(0,100)
      .build();

    // TIEBREAKER 1 -- combined final score of both teams
    let helpTextOne = `Combined points between ${tiebreakerMatchup}${config.overUnderInclude && tiebreakerOverUnder > 0 ? ' (betting line: ' + tiebreakerOverUnder + ')' : ''}`;
    form.addTextItem()
      .setTitle('Tiebreaker 1')
      .setHelpText(helpTextOne)
      .setRequired(true)
      .setValidation(combinedValidation);

    // TIEBREAKER 2 -- points scored by the winning team of that same game
    let helpTextTwo = `Points scored by the WINNING team in ${tiebreakerMatchup} (only used if Tiebreaker 1 is also tied)`;
    form.addTextItem()
      .setTitle('Tiebreaker 2')
      .setHelpText(helpTextTwo)
      .setRequired(true)
      .setValidation(winnerValidation);

    ss.toast(`Created both tiebreaker questions for ${tiebreakerGame.awayTeam} @ ${tiebreakerGame.homeTeam}`,`⚖️ TIEBREAKERS CREATED`);
    Logger.log(`⚖️ Two tiebreaker questions created for the ${mnf ? 'late MNF game' : 'last game of the week'}: ${tiebreakerMatchup}`);
  }
  if(!config.commentsExclude) { // Excludes comment question if comments are disabled
    form.addTextItem()
      .setTitle('Comments')
      .setHelpText('Passing thoughts...');
    ss.toast(`Added comment box for pick 'ems`,`✍ COMMENT BOX CREATED`);
    Logger.log(`✍ Comments field added for pick 'ems`);
  }
}

/**
 * Builds the page for a new user to sign up.
 * This function now intelligently adds Survivor and/or Eliminator questions
 * directly to the same page if it's the first week of the contest.
 *
 * @param {Form} form The Google Form object to add items to.
 * @param {Object} config The main configuration object for the pool.
 * @param {Object} gamePlan The game plan for the current week.
 * @returns {PageBreakItem} The created page break item for navigation.
 */
function buildNewUserPage(ss, form, config, gamePlan, survivor, eliminator, survivorStart, eliminatorStart) {
  // 1. Create the page break and set its final destination.
  const newUserPage = form.addPageBreakItem();
  newUserPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
  
  // --- [THE NEW LOGIC] ---
  // 3. Conditionally add contest questions directly to this page.  
  contests = 0;
  if ((survivor && survivorStart) || (eliminator && eliminatorStart)) {
    form.addSectionHeaderItem().setTitle(`🏆 Contest Pick${(survivor && survivorStart) && (eliminator && eliminatorStart) ? 's' : ''}`);
  }
  if (survivor) {
    if (survivorStart) {
      Logger.log(`👑 NEW USER: Eligible for survivor start week (${gamePlan.week}), adding queston...`);
      let allSurvivorTeamsForWeek = buildTeamList(gamePlan, config, config.survivorAts);
      addContestQuestion(form, 'survivor', {}, config.survivorAts, config.survivorStartWeek, allSurvivorTeamsForWeek, survivorStart); // Pass empty member object
      ss.toast(`Generated new user survivor question`,`👑 NEW USER SURVIVOR`);
      contests++;
    } else {
      Logger.log(`👑 NEW USER: Ineligible for survivor start week (${gamePlan.week})`);
    }
  } else {
    Logger.log(`👑 NEW USER: No eliminator pool present`);
  }

  if (eliminator) {
    if (eliminatorStart) {
      Logger.log(`💀 NEW USER: Eligible for eliminator start week (${gamePlan.week}), adding queston...`);
      let allEliminatorTeamsForWeek = buildTeamList(gamePlan, config, config.eliminatorAts);
      addContestQuestion(form, 'eliminator', {}, config.eliminatorAts, config.eliminatorStartWeek, allEliminatorTeamsForWeek, eliminatorStart); // Pass empty member object
      ss.toast(`Generated new user eliminator question`,`💀 NEW USER ELIMINATOR`);
      contests++;
    } else {
      Logger.log(`💀 NEW USER: Ineligible for eliminator start week (${gamePlan.week})`);
    }
  } else {
    Logger.log(`💀 NEW USER: No eliminator pool present`);
  }

  if (contests > 0) {
    form.addSectionHeaderItem().setTitle('✏️ Name Entry');
    newUserPage.setTitle('New User Contest Entry & Signup')
  } else {
    newUserPage.setTitle('New User Signup')
  }

  form.addTextItem()
    .setTitle('Enter Your Name')
    .setHelpText('Please enter your name as it will appear in the pool.')
    .setRequired(true)
    .setValidation(nameValidation);
  // 4. Return the page break item so the main function can use it for navigation.
  return newUserPage;
}

/**
 * Takes all games within gameplan and populates an array with the team abbreviations, including emojis if enabled
 * 
 * @param {object} all data submitted by form creation panel
 * @param {object} league data with necessary booleans/values
 * @param {boolean} optional value to determin if spreads should be added
 * @returns {array} all teams selected for the given week to provide in a survivor/eliminator drop-down with ATS
 */
function buildTeamList(gamePlan, config, isAts) {
  isAts = isAts || (config.survivorAts || config.elminatorAts); // Fallback
  return gamePlan.games.flatMap(game => {
    const awayEmojis = `${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.awayTeam].mascot : ''}`;
    const homeEmojis = `${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.homeTeam].mascot : ''}`;
    if (!isAts || game.spread === 'PK' || game.spread == 0) {
      return [
        `${game.awayTeam}${awayEmojis}`,
        `${game.homeTeam}${homeEmojis}`];
    }
    // Extract the numeric value and determine which team is favored
    const spreadMatch = game.spread.match(/([A-Z]+)\s*(-?\d+\.?\d*)/);
    if (!spreadMatch) return [`${game.awayTeam}${awayEmojis}`,`${game.homeTeam}${homeEmojis}`]; // Fallback
    
    const [,favoriteTeam, spreadValue] = spreadMatch;
    const numericSpread = parseFloat(spreadValue);
    
    // Determine spreads for away and home teams
    const awaySpread = favoriteTeam === game.awayTeam ? numericSpread : Math.abs(numericSpread);
    const homeSpread = favoriteTeam === game.homeTeam ? numericSpread : Math.abs(numericSpread);
    
    return [
      `${game.awayTeam} ${awaySpread > 0 ? '+' : ''}${awaySpread}${awayEmojis}`,
      `${game.homeTeam} ${homeSpread > 0 ? '+' : ''}${homeSpread}${homeEmojis}`
    ];
  }).sort((a, b) => a.charAt(0).localeCompare(b.charAt(0)));
}

/**
 * Combines emoji or square entries to present user with visual of remaining lives
 * 
 * @param {boolean} state of whether emojis are present or not
 * @param {integer} remaining lives per user
 * @param {integer} total lives by default
 * @returns {string} character string representing lives left and those have been lost (red dot or empty square)
 */
function createLivesString(remaining, total) {
  return '🟢'.repeat(remaining)+'⚫'.repeat(total - remaining);
}

/**
 * Checks for existence of a file based on given ID
 * 
 * @param {string} file ID to check
 * @returns {boolean} true if file exists, false if not
 */
function checkFileExists(fileId) {
  // Check if fileId is valid (not null, undefined, or empty string)
  if (!fileId || typeof fileId !== 'string' || fileId.trim() === '') {
    return false;
  }  
  try {
    // DriveApp.getFileById works for both files AND folders
    const file = DriveApp.getFileById(fileId);
    if (file && !file.isTrashed()) {
      return true;
    } else {
      return false;
    }
  } catch (err) {
    Logger.log(`❗ "checkFileExists" did not locate a file for provided file ID ${fileId}`);
    return false; // Exception means file/folder doesn't exist or isn't accessible
  }
}

/**
 * Checks for existence of a folder based on given ID
 * 
 * @param {string} folder ID to check
 * @returns {boolean} true if folder exists, false if not
 */
function checkFolderExists(folderId) {
  if (!folderId || typeof folderId !== 'string' || folderId.trim() === '') {
    return false;
  }
  try {
    const folder = DriveApp.getFolderById(folderId);
    if (folder && !folder.isTrashed()) {
      return true;
    } else {
      return false;
    }
  } catch (err) {
    Logger.log(`❗ "checkFolderExists" did not locate a folder for provided folder ID ${folderId}`);
    return false;
  }
}

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// Convert 24 hour entry into AM/PM string with minutes
function formatTime(hour,minute) {
  const suffix = hour >= 12 ? 'PM' : 'AM';
  return `${hour > 12 ? hour - 12 : hour}:${minute < 10 ? '0' + minute : minute} ${suffix}`;
}

/**
 * Updates the 'Schedule' sheet using the final, authoritative
 * gamePlan object sent from the form creator panel.
 *
 * @param {Object} gamePlan The complete game plan, including user-overridden spread/OU data.
 */
function updateScheduleData(ss,gamePlan) {
  // If the gamePlan has no games, there's nothing to update.
  ss = fetchSpreadsheet(ss) || SpreadsheetApp.getActiveSpreadsheet();

  if (!gamePlan || !gamePlan.games || gamePlan.games.length === 0) {
    return;
  }
  
  const week = gamePlan.week;
  Logger.log(`Applying data overrides from gamePlan for Week ${week}.`);
  
  const sheet = ss.getSheetByName(LEAGUE); // e.g., 'NFL'
  if (!sheet) throw new Error(`The '${LEAGUE}' sheet could not be found.`);

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  
  // Find the column indexes for all relevant fields.
  const weekCol = headers.indexOf('week');
  const awayCol = headers.indexOf('awayTeam');
  const homeCol = headers.indexOf('homeTeam');
  const spreadCol = headers.indexOf('spread');
  const overUnderCol = headers.indexOf('overUnder');
  const timeFetchedCol = headers.indexOf('timeFetched');
  const autoFetchedCol = headers.indexOf('spreadAutoFetched');

  // Create a fast lookup map of the sheet's data.
  // Key: "AWAY@HOME", Value: the row's index in the 'data' array.
  const sheetDataMap = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[weekCol] == week) {
      const gameId = `${row[awayCol]}@${row[homeCol]}`;
      sheetDataMap.set(gameId, i); // Store the index (e.g., 1 for the second row)
    }
  }

  let changesMade = 0;
  
  // Loop through the games from the INCOMING GAMEPLAN. This is our source of truth.
  gamePlan.games.forEach(game => {
    const gameId = `${game.awayTeam}@${game.homeTeam}`;
    
    // Find the corresponding row in our sheet data using the map.
    const rowIndex = sheetDataMap.get(gameId);
    
    if (rowIndex !== undefined) {
      // We found a match! Now, update the data in our in-memory 'data' array.
      
      // Update the spread and Over/Under with the final values from the gamePlan.
      data[rowIndex][spreadCol] = game.spread;
      data[rowIndex][overUnderCol] = game.overUnder;
      
      // As requested, update the timestamp and auto-fetch status.
      data[rowIndex][timeFetchedCol] = new Date();
      if (autoFetchedCol > -1) {
        data[rowIndex][autoFetchedCol] = 0; // Mark as a manual entry/override
      }
      changesMade++;
    }
  });

  // If any changes were made, write the entire updated data array back to the sheet.
  if (changesMade > 0) {
    dataRange.setValues(data);
    Logger.log(`Successfully updated the Schedule sheet with ${changesMade} game overrides.`);
    ss.toast(`Successfully updated the Schedule sheet with ${changesMade} game overrides.`,`📝 ${changesMade} SPREAD UPDATES`)
  }
}

/**
 * The main function to launch the "Auto-Fetch Settings" panel.
 * This should be called from a menu item.
 */
function showAutoFetchPanel() {
  const html = HtmlService.createHtmlOutputFromFile('triggerPanel')
      .setWidth(450)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Auto-Fetch Settings');
}

/**
 * Fetches the details of the existing weekly fetch trigger, if it exists.
 * This is called by the panel when it loads.
 * @returns {Object|null} An object with { day, hour } or null if no trigger is found.
 */
function getWeeklyFetchTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  let triggerInfo = { configured : false };
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'runWeeklyFetch') {
      // Unfortunately, Apps Script doesn't let us read the day/hour directly.
      // We must store this information in properties when we create the trigger.
      triggerInfo = JSON.parse(PropertiesService.getDocumentProperties().getProperty('weeklyFetchTriggerInfo'));
      if (triggerInfo) {
        triggerInfo.configured = true;
        Logger.log(JSON.stringify(triggerInfo))
        return triggerInfo;
      }
    }
  }
  return triggerInfo; // No trigger found
}

/**
 * Creates or updates the weekly time-based trigger.
 * @param {Object} data An object with { day, hour }.
 */
function setWeeklyFetchTrigger(data) {
  const { day, hour } = data;

  // First, delete any existing trigger to ensure there's only one.
  deleteWeeklyFetchTrigger();

  // Create the new trigger to run the 'runWeeklyFetch' wrapper function.
  ScriptApp.newTrigger('runWeeklyFetch')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay[day.toUpperCase()])
    .atHour(hour)
    .nearMinute(15) // Run sometime around the 15-minute mark to distribute load
    .create();

  // Store the settings so we can display them to the user later.
  const triggerInfo = { day: day, hour: hour };
  PropertiesService.getDocumentProperties().setProperty('weeklyFetchTriggerInfo', JSON.stringify(triggerInfo));

  SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Auto-fetch scheduled for every ${day} around ${hour}:00.`);
  return { success: true };
}

/**
 * Deletes the weekly fetch trigger.
 */
function deleteWeeklyFetchTrigger() {
  let wasDeleted = false;
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === 'runWeeklyFetch') {
      ScriptApp.deleteTrigger(trigger);
      wasDeleted = true;
    }
  }
  
  // Also clear the stored info
  PropertiesService.getDocumentProperties().deleteProperty('weeklyFetchTriggerInfo');
  
  if (wasDeleted) {
    SpreadsheetApp.getActiveSpreadsheet().toast('❌ Auto-fetch has been disabled.');
  }
  return { success: true };
}



/**
 * The main data-gathering function for the Import Picks panel.
 * This is called by the client-side script on load.
 */
function getFormImportData(week,auto) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const formsData = JSON.parse(docProps.getProperty('forms'));
    let memberData = JSON.parse(docProps.getProperty('members')) || { memberOrder: [], members: {}};
    week = week || 1;
    if (formsData) {
      week = week || Math.max(...Object.keys(formsData).map(key => parseInt(key)));
      Logger.log(week);
    }
    // Run sync first to get the latest respondent metadata.
    let syncResult;
    try {
      syncResult = syncFormResponses(week);
    } catch (err) {
      Logger.log(`Week ${week} not created yet or unavailable, moving on to panel loading.`)
    }
    // Fetch the clean, final data needed for the panel.
    const config = JSON.parse(docProps.getProperty('configuration'));
    
    const allCreatedWeeks = Object.keys(formsData).map(Number).sort((a,b) => a-b);
    week = auto ? (allCreatedWeeks[allCreatedWeeks.length - 1] || fetchWeek(null, true)) : (week || allCreatedWeeks[allCreatedWeeks.length - 1] || fetchWeek(null, true));
    let gamePlan = formsData[week]?.gamePlan;

    if (!gamePlan) {
      Logger.log(`⚠️ No form found for week ${week}`);
      Logger.log(`🔎 attempting to look for most recent week created...`);
      week = allCreatedWeeks[allCreatedWeeks.length - 1] || fetchWeek(null, true);
      gamePlan = formsData[week]?.gamePlan;
      if (!gamePlan) {
        Logger.log(`❓ It appears no gamePlans exist yet for forms. Are you sure you've created a form already?`)
        Logger.log(`⚠️ Error in getFormImportData: ${err.stack}`);
        throw new Error(`Failed to prepare for import: ${err.stack}`);
      }
    }

    // Determine which games are upcoming.
    const matchups = getInvalidPickMatchups();
    
    // Determine if a partial import should be offered.
    const allMembersResponded = syncResult?.totalRespondents ? syncResult.totalRespondents === memberData.memberOrder.length : false;
    const isMembershipLocked = config.membershipLocked;
    const offerPartialImport = !(allMembersResponded && isMembershipLocked);

    // Bundle all data and return it to the client.
    return {
      week: week,
      allCreatedWeeks: allCreatedWeeks || [],
      newMembers: formsData[week]?.newMembers || [],
      respondentIds: formsData[week]?.respondents || [],
      allMemberIds: memberData.memberOrder,
      members: memberData.members,
      gamePlanGames: gamePlan.games,
      startedGames: matchups,
      offerPartialImport: offerPartialImport
    };
  } catch (err) {
    Logger.log(`⚠️ Error in getFormImportData: ${err.stack}`);
    // Re-throw the error so the client's onFailure handler gets it.
    throw new Error(`⚠️ Failed to prepare for import: ${err.stack}`);
  }
}
/**
 * [MODIFIED] This function now only launches the HTML file.
 */
function launchFormImport() {
  const html = HtmlService.createHtmlOutputFromFile('formImport')
      .setWidth(600)
      .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Weekly Picks');
}

/**
 * Imports all processed picks into the correct weekly
 * pick'em sheet, survivor sheet, and eliminator sheet.
 *
 * @param {number} week The week to import.
 * @param {boolean} importOnlyStartedGames If true, only imports picks for games that have already started.
 */
function executePickImport(week, importOnlyStartedGames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // --- 1. Fetch All Necessary Data ---
  
  const docProps = PropertiesService.getDocumentProperties();
  let config = JSON.parse(docProps.getProperty('configuration')) || {};
  const memberData = JSON.parse(docProps.getProperty('members')) || {};
  let formsData = JSON.parse(docProps.getProperty('forms')) || {};
  const databaseSheet = getDatabaseSheet();
  const responseSheet = databaseSheet.getSheetByName(`WK${week}`);
  
  // Parse the latest, de-duplicated picks from the response sheet.
  const parsedPicks = parseAllPicksFromSheet(responseSheet, memberData);
  // --- 2. Handle Pick'em Sheet Population ---
  if (config.pickemsInclude) {
    try {
      const weeklySheetName = `${weeklySheetPrefix}${week}`;
      let sheet = ss.getSheetByName(weeklySheetName);
      if (!sheet || !ss.getRangeByName(`NAMES_${week}`)) {
        Logger.log(`🔁 No weekly sheet exists for week ${week}, creating one now...`);
        ss.toast(`Creating weekly sheet for week ${week}.`,'🔁 CREATING...');
        let displayEmpty = true;
        if (config?.hideNonParticipants) displayEmpty = !config.hideNonParticipants;
        sheet = weeklySheet(ss,week,config,formsData,memberData,displayEmpty);
      }
      // --- Create Lookup Maps ---
      // a) Member Name -> Row Index Map (Unchanged)
      let memberNameRange;
      try { 
        memberNameRange = ss.getRangeByName(`NAMES_${week}`);
      } catch (err) {
        Logger.log(`❗ Error getting named range for WK${week} names: ${err.stack}`);
        let memberNameRangeStart = 3, memberNameRangeEnd = 0, index = memberNameRangeStart;
        while (memberNameRangeEnd < memberNameRangeStart) {
          if (sheet.getRange(index,1).getValue() == 'PREFERRED') {
            memberNameRangeEnd = index;
          }
          if (index > 100) {
            throw new Error(`⚠️ Unable to locate NAMES range for week ${week} weekly sheet even with fallback. Failed to import.`);
          }
          index++;
        }
        memberNameRange = sheet.getRange(memberNameRangeStart,1,memberNameRangeEnd - memberNameRangeStart,1);
        Logger.log(`✅ Successfully pulled NAMES range by lookup value despite error being thrown.`);
      }
      if (!memberNameRange) throw new Error(`⚠️ Unable to locate NAMES range for week ${week} weekly sheet. Failed to import.`);
      const memberNames = memberNameRange.getValues().flat();
      const memberNameToRowMap = new Map(memberNames.map((name, index) => [name, index]));

      // --- [THE NEW LOGIC] Team-Pair Matching ---
      // b) Matchup -> Column Index Map
      const matchupRange = ss.getRangeByName(`${LEAGUE}_${week}`);
      if (!matchupRange) throw new Error(`Named range '${LEAGUE}_${week}' not found.`);
      
      const matchupHeaders = matchupRange.getValues()[0];
      const matchupToColMap = new Map();
      matchupHeaders.forEach((header, index) => {
        const teams = header.toString().match(/[A-Z]{2,3}/g);
        if (teams && teams.length === 2) {
          const teamKey = teams.sort().join('-'); // e.g., "BUF-MIA"
          matchupToColMap.set(teamKey, index);
        }
      });
      // --- Prepare Data for Writing (Unchanged) ---
      const picksRange = ss.getRangeByName(`${LEAGUE}_PICKS_${week}`);
      let tiebreakerRange, tiebreakers, tiebreaker2Range, tiebreakers2;
      if (config.tiebreakerInclude) {
        tiebreakerRange = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}`);
        if (tiebreakerRange) tiebreakers = tiebreakerRange.getValues();
        tiebreaker2Range = ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}`);
        if (tiebreaker2Range) tiebreakers2 = tiebreaker2Range.getValues();
      }
      let commentRange, comments;
      if (!config.commentsExclude) {
        commentRange = ss.getRangeByName(`COMMENTS_${week}`);
        if (commentRange) comments = commentRange.getValues();
      }
      
      if (!picksRange) throw new Error(`Named range '${LEAGUE}_PICKS_${week}' not found.`);
      const picksData = picksRange.getValues();
      const gamePlan = formsData[week]?.gamePlan;
      let startedGames = new Set(getStartedGames());
  
      // --- 3. Loop Through Parsed Picks and Populate the 2D Array ---
      for (const memberId in parsedPicks) {
        const member = memberData.members[memberId];
        const picks = parsedPicks[memberId];
        const rowIndex = memberNameToRowMap.get(member.name);
        if (rowIndex === undefined) continue;
        
        for (const question in picks.pickem) {
          const pick = picks.pickem[question];

          // Find the original game from the gamePlan to get the team pair.
          const game = gamePlan.games.find(g => question.includes(g.awayTeamName) && question.includes(g.homeTeamName));
          if (game) {
            // Apply the import filter first for efficiency
            const matchupShortName = `${game.awayTeam} @ ${game.homeTeam}`;
            const matchupVsName = `${game.awayTeam} VS ${game.homeTeam}`;
            
            if (importOnlyStartedGames && !startedGames.has(matchupShortName) && !startedGames.has(matchupVsName)) {
              continue; // Skip if it's an upcoming game
            }
            
            // --- [THE NEW LOGIC] Find the column using the team-pair key ---
            const teamKey = [game.awayTeam, game.homeTeam].sort().join('-');
            const colIndex = matchupToColMap.get(teamKey);

            if (colIndex !== undefined) {
              // Check if the member's pick is actually one of the teams in the matchup
              if (pick === game.awayTeam || pick === game.homeTeam) {
                picksData[rowIndex][colIndex] = pick;
              }
            }
          }
        }
        if (picks.tiebreaker && tiebreakers) tiebreakers[rowIndex][0] = picks.tiebreaker;
        if (picks.tiebreaker2 && tiebreakers2) tiebreakers2[rowIndex][0] = picks.tiebreaker2;
        if (picks.comments) comments[rowIndex][0] = picks.comments;
      }
      
      // --- 4. Write Data Back to the Sheet (Unchanged) ---
      picksRange.setValues(picksData);
      if (!importOnlyStartedGames && config.tiebreakerInclude && tiebreakerRange && tiebreakers) tiebreakerRange.setValues(tiebreakers);
      if (!importOnlyStartedGames && config.tiebreakerInclude && tiebreaker2Range && tiebreakers2) tiebreaker2Range.setValues(tiebreakers2);
      if (!config.commentsExclude) commentRange.setValues(comments);
      const text = `Successfully imported Pick 'Em data into week '${week}' sheet.`;
      Logger.log(`✅ ${text}`);
      ss.toast(text,`✅ PICK 'EMS IMPORTED`)
      
      formsData[week].imported = true;
      saveProperties('forms',formsData);

      if (!config.initialized) {
        const ui = fetchUi();
        let prompt = ui.alert(`Season-Long Tracking Sheets Creation`,'Would you like to create all additional tracking sheets now?\n\nThis can be done later via the "Picks" > "Utilities" menu.', ui.ButtonSet.YES_NO);
        if (prompt == "YES") {
          try {
            setupSheets();
          } catch (err) {
            ss.toast('Issue creating setup sheets, run again from the utilities menu');
            Logger.logger('Issue creating setup sheets | ERROR: ' + err.stack);
          }
          ss.toast(`Successfully configured all other sheets.`, '✅ SETUP SHEETS SUCCESS');
          config.initialized = true;
          saveProperties('configuration',config);
        } else {
          ss.toast(`Declined setup of all other sheets, try again via the "Picks" > "Utilities" menu later.`,`❎ NO SETUP SHEETS`);
        }
      }
    } catch (err) {
      const text = `Failed to import Pick 'Em data into week '${week}' sheet.`;
      Logger.log(`${text} | ERROR: ${err.stack}`);
      ss.toast(text,`❗  PICK 'EMS FAILED`)

    }
  }
  if (!importOnlyStartedGames) {
    let survInclude = config.survivorInclude && week >= config.survivorStartWeek;
    const elimInclude = config.eliminatorInclude && week >= config.eliminatorStartWeek;
    // --- 5. Populate Survivor and Eliminator Sheets ---
    if (survInclude) populateSurvElimSheet(ss, parsedPicks, memberData, config, formsData[week]?.gamePlan, week, 'survivor');
    if (elimInclude) populateSurvElimSheet(ss, parsedPicks, memberData, config, formsData[week]?.gamePlan, week, 'eliminator');
    // Record picks from applicable survivor/eliminator to the JSON members object
    if (survInclude || elimInclude) recordSurvElimResponses(parsedPicks, memberData, week, survInclude, elimInclude);

  } else {
    const title = ((config.survivorInclude && week >= config.survivorStartWeek) && (config.eliminatorInclude && week >= config.eliminatorStartWeek)) ? `NO SURVIVOR/ELMINATOR YET` : (config.survivorInclude && week >= config.survivorStartWeek) ? `NO SURVIVOR YET` : `NO ELIMINATOR YET`;
    const notification = ((config.survivorInclude && week >= config.survivorStartWeek) && (config.eliminatorInclude && week >= config.eliminatorStartWeek)) ?
      `Survivor and Eliminator not imported: user declined to import all matchups.` : (config.survivorInclude && week >= config.survivorStartWeek) ? 
      `Survivor not imported: user declined to import all matchups.` : (config.eliminatorInclude && week >= config.eliminatorStartWeek) ? `Eliminator not imported: user declined to import all matchups.` : `Currently no Survivor or Eliminator pool to import`;
    Logger.log(`❎ ${notification}`);
    ss.toast(notification,`❎ ${title}`);    
  }
  
  // Updates the Outcomes sheet to reflect the games that were actually being evaluated by the form, resets conditional formatting and data validation rules, then checks if Pick 'Ems present, whether any values were in place on the Outcomes sheet already and replaces them after otherwise putting a connection in place back to the weekly sheet
  try {
    outcomesSheetUpdate(ss,week,config,formsData[week].gamePlan)
    const text = `✅ Successfully updated the OUTCOMES sheet input ranges for week '${week}' range.`;
    Logger.log(text);
    ss.toast(text,`OUTCOMES SHEET UPDATED`);
  } catch (err) {
    const text = `❗ Failed to update the OUTCOMES sheet input ranges for week '${week}' range.`;
    Logger.log(text + ' | ERROR: ' + err.stack);
    ss.toast(text,`OUTCOMES NOT UPDATED`);
  }

  // --- 6. Finalize and Save ---
  formsData[week].imported = true;
  saveProperties('forms', formsData);

  return { success: true, message: `✅ Picks for week ${week} have been successfully imported!` };
}

/**
 * Updates the memberData object with Survivor and Eliminator picks for a specific week
 * and saves the updated object to Document Properties.
 * 
 * @param {Object} parsedPicks - The object containing picks (keyed by memberId).
 * @param {Object} memberData - The master members object.
 * @param {number|string} week - The current NFL week number.
 */
function recordSurvElimResponses(parsedPicks, memberData, week, survInclude, elimInclude) {
  // Safeguards for testing or absent entries
  const verbose = !parsedPicks || !memberData || !week;
  week = week || 1;
  memberData = memberData || JSON.parse(PropertiesService.getDocumentProperties().getProperty('members'));
  parsedPicks = parsedPicks || parseAllPicksFromSheet(getDatabaseSheet().getSheetByName(`WK${week}`),memberData)
  const weekIdx = parseInt(week) - 1; // Convert Week 1 to Index 0
  let updateCount = 0;

  if (verbose) {
    Logger.log(`🧾 Verbose Mode: Displaying Week ${week} Responses:`);
    memberData.memberOrder.forEach(id => {
      const member = memberData.members[id];
      const picks = parsedPicks[id]; // Matches by the ID key (e.g., id_8VS4MCM966)

      if (picks) {
        const s = picks.survivor || "---";
        const e = picks.eliminator || "---";
        if (verbose || (survInclude && elimInclude)) {
          Logger.log(`👤 ${member.name}: [Survivor: ${s}] [Eliminator: ${e}]`);
        } else if (survInclude) {
          Logger.log(`👤 ${member.name}: [Survivor: ${s}]`);
        } else {
          Logger.log(`👤 ${member.name}: [Eliminator: ${e}]`);
        }
      } else {
        // Optional: Log if a member didn't submit anything at all
        Logger.log(`👤 ${member.name}: No response received.`);
      }
    });
  }

  try {
    // 1. Iterate through every member present in the current picks batch
    for (const memberId in parsedPicks) {
      const member = memberData.members[memberId];
      
      // Skip if the member doesn't exist in our master record
      if (!member) {
        Logger.log(`⚠️ Warning: Member ID ${memberId} found in picks but not in master member list.`);
        continue;
      }

      const picks = parsedPicks[memberId];

      // 2. Handle Survivor Pick (sP)
      if (picks.survivor) {
        if (!member.sP) member.sP = [];
        // Fill gaps with null if necessary
        while (member.sP.length < weekIdx) {
          member.sP.push(null);
        }
        member.sP[weekIdx] = picks.survivor;
      }

      // 3. Handle Eliminator Pick (eP)
      if (picks.eliminator) {
        if (!member.eP) member.eP = [];
        // Fill gaps with null if necessary
        while (member.eP.length < weekIdx) {
          member.eP.push(null);
        }
        member.eP[weekIdx] = picks.eliminator;
      }

      updateCount++;
    }

    // 4. Save the modified object back to PropertiesService
    saveProperties('members', memberData);
    
    if (verbose) {
      Logger.log(`🔎 Displaying document properties...`)
      viewDocumentProperties;
    }

    Logger.log(`✅ Successfully logged contest picks for ${updateCount} members for Week ${week}.`);
    return true;

  } catch (err) {
    Logger.log(`❌ Error in logWeeklyContestPicks: ${err.stack}`);
    return false;
  }
}


// You will also need this small helper function (modified from a previous version)
function getStartedGames() {
  try {
    const response = espnFetch(SCOREBOARD);
    const data = JSON.parse(response.getContentText()).events;
    // An empty set is fine, as it is used to filter out games that are past kickoff and will not be imported
    const startedGames = new Set();
    for (const event of data) {
      if (event.status.type.state !== "pre") { // "in" or "post"
        startedGames.add(event.shortName);
      }
    }
    return startedGames;
  } catch (e) {
    return new Set();
  }
}

const filterMatchups = (matchupToColMap, startedGames) => {
  const startedTeams = new Set();
  
  // Extract all teams from started games
  startedGames.forEach(game => {
    const teams = game.toString().match(/([A-Z]{2,3})/g);
    if (teams) {
      teams.forEach(team => startedTeams.add(team));
    }
  });
  
  // Group by column index and filter
  const colGroups = {};
  Object.entries(matchupToColMap).forEach(([team, col]) => {
    if (!colGroups[col]) colGroups[col] = [];
    colGroups[col].push(team);
  });
  
  const result = {};
  Object.entries(colGroups).forEach(([col, teams]) => {
    if (teams.every(team => startedTeams.has(team))) {
      teams.forEach(team => result[team] = parseInt(col));
    }
  });
  
  return result;
};


/**
 * Intelligently shows or hides the "Margin" columns on the OUTCOMES sheet
 * based on whether any ATS mode is active in the configuration.
 *
 * @param {Object} [config] Optional: The configuration object. If not provided, it will be fetched.
 */
function updateOutcomeSheetVisibility(config) {
  try {
    // 1. Receive or Fetch Config
    if (!config) {
      const docProps = PropertiesService.getDocumentProperties();
      config = JSON.parse(docProps.getProperty('configuration') || '{}');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const outcomesSheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);
    if (!outcomesSheet) {
      Logger.log(`❗ NO OUTCOMES SHEET FOUND: Sheet not found. Cannot update column visibility.`);
      return; // Exit gracefully if the sheet doesn't exist.
    }

    // 2. Determine the required state
    const isAnyAtsActive = config.pickemsAts || config.survivorAts || config.eliminatorAts;
    
    // 3. Check the current state of the first margin column to avoid unnecessary actions.
    const firstMarginColumn = 2; // Column B
    const isCurrentlyVisible = !outcomesSheet.isColumnHiddenByUser(firstMarginColumn);

    // 4. Act only if the current state doesn't match the required state.
    if (isAnyAtsActive && !isCurrentlyVisible) {
      // ATS is ON, but columns are HIDDEN -> UNHIDE them.
      Logger.log(`✅ MARGIN COLUMNS VISIBLE: Showing ATS Margin columns...`);
      // Loop through all even columns and unhide them individually.
      // Unfortunately, Sheets API does not have an "unhideColumns" for multiple non-contiguous ranges.
      for (let i = firstMarginColumn; i <= outcomesSheet.getMaxColumns(); i += 2) {
        outcomesSheet.unhideColumn(outcomesSheet.getRange(1, i));
      }
      outcomesSheet.getRange(ss.getRangeByName(`${LEAGUE}_OUTCOMES_1`).getRow()-2,1,1,outcomesSheet.getMaxColumns()).setFontSize(8);
      outcomesSheet.getRange(ss.getRangeByName(`${LEAGUE}_OUTCOMES_1`).getRow()-1,1,1,outcomesSheet.getMaxColumns()).setFontSize(10);
      outcomesSheet.setColumnWidths(1,outcomesSheet.getMaxColumns(),50);
    } else if (!isAnyAtsActive && isCurrentlyVisible) {
      // ATS is OFF, but columns are VISIBLE -> HIDE them.
      Logger.log(`✅ MARGIN COLUMNS HIDDEN: Hiding ATS Margin columns...`);
      // It's more efficient to hide columns in batches if possible, but looping is reliable.
      for (let i = firstMarginColumn; i <= outcomesSheet.getMaxColumns(); i += 2) {
        outcomesSheet.hideColumns(i);
      }
      outcomesSheet.getRange(ss.getRangeByName(`${LEAGUE}_OUTCOMES_1`).getRow()-2,1,1,outcomesSheet.getMaxColumns()).setFontSize(7);
      outcomesSheet.getRange(ss.getRangeByName(`${LEAGUE}_OUTCOMES_1`).getRow()-1,1,1,outcomesSheet.getMaxColumns()).setFontSize(8);
      outcomesSheet.setColumnWidths(1,outcomesSheet.getMaxColumns(),60);
    }
    // If state is already correct (e.g., ATS is on and columns are visible), do nothing.

  } catch (err) {
    Logger.log(`❗ ERROR HIDING COLUMNS: Error updating OUTCOMES sheet visibility: ${err.stack}`);
  }
}




/**
 * Reads a raw form response sheet, de-duplicates to "last
 * submission wins", and parses all pick types into a clean, structured object.
 * This is a read-only operation and does not modify any properties.
 *
 * @param {Sheet} sheet The Google Sheet object for a specific week's responses (e.g., 'WK1').
 * @param {Object} memberData The complete, current 'members' object.
 * @returns {Object} A "picks cache" object, where keys are member IDs and values
 *                   are objects containing all of that member's final picks.
 *                   e.g., { "id_123": { pickem: {...}, survivor: "BAL", ... } }
 */
function parseAllPicksFromSheet(sheet, memberData) {
  if (!sheet) {
    Logger.log(`⭕ "parseAllPicksFromSheet" was called with a null sheet. Returning empty object.`);
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log(`No responses found in sheet '${sheet.getName()}' to parse.`);
    return {};
  }

  const headers = data.shift();
  
  // --- Find critical column indexes using helper ---
  const { nameCol, newNameCol } = findNameColumns(headers);
  if (nameCol === -1 && newNameCol === -1) {
    Logger.log(`⚠️ CRITICAL: Could not find a 'Select Your Name' or 'Enter Your Name' entry from responses. Cannot parse picks.`);
    return {};
  } else if (nameCol === -1) {
    Logger.log(`❕ NOTE: Form appears to be in open enrollment mode for this week.`);
  }

  // --- 1. De-duplicate to "Last Submission Wins" ---
  const latestSubmissions = {};
  const newUserAnswerRegex = /new user/i;
  data.forEach(row => {
    const name = (nameCol === -1 || newUserAnswerRegex.test(row[nameCol]))
      ? row[newNameCol] 
      : row[nameCol];
    if (name && name.trim() !== '') {
      latestSubmissions[name.trim().toLowerCase()] = row;
    }
  });
  const finalResponseRows = Object.values(latestSubmissions);

  // --- 2. Create a lookup map for name -> ID ---
  const nameToIdMap = {};
  for (const id in memberData.members) {
    nameToIdMap[memberData.members[id].name.toLowerCase()] = id;
  }

  // --- 3. [THE NEW ENGINE] Intelligently Parse and Clean ALL Picks ---
  const weeklyPicksCache = {};
  // This regex will find the first 2-3 letter capital word in a string.
  const teamAbbrRegex = /[A-Z]{2,3}/; 
  const emojiAndSpecialCharsRegex = /[^A-Z0-9\s\-.()+]/gi;
  
  const survivorRegex = /survivor/i;
  const eliminatorRegex = /eliminator/i;
  const tiebreaker2Regex = /tiebreaker\s*2/i; // must be tested before the generic one below
  const tiebreakerRegex = /tiebreaker/i;
  const commentsRegex = /comments/i;
  const pickemRegex = / at /i;

  finalResponseRows.forEach(row => {
    const name = (nameCol === -1 || newUserAnswerRegex.test(row[nameCol])) ? row[newNameCol] : row[nameCol];
    const memberId = nameToIdMap[name.trim().toLowerCase()];
    
    // If we can't map the submission to an existing member ID, we skip it.
    // (New members would have been added in the 'sync' step before this is called).
    if (!memberId) {
        Logger.log(`Skipping picks for "${name}" as they could not be mapped to a member ID.`);
        return; // 'continue' in a forEach loop
    }

    const userPicks = {
      pickem: {},
      survivor: null,
      eliminator: null,
      tiebreaker: null,
      tiebreaker2: null,
      comments: ''
    };

    headers.forEach((header, index) => {
      let answer = row[index];
      if (answer === '' || answer === null || answer === undefined) return;
      
      // Convert to string for consistent processing
      answer = answer.toString().replace(emojiAndSpecialCharsRegex, '').trim();

      const question = header;
      if (commentsRegex.test(question)) {
        userPicks.comments = answer;
      } else {
        answer = answer.toString().replace(emojiAndSpecialCharsRegex, '').trim();
        if (survivorRegex.test(question) && answer) {
          userPicks.survivor = answer;
        } else if (eliminatorRegex.test(question) && answer) {
          userPicks.eliminator = answer;
        } else if (tiebreaker2Regex.test(question)) {
          userPicks.tiebreaker2 = answer;
        } else if (tiebreakerRegex.test(question)) {
          userPicks.tiebreaker = answer;
        } else if (pickemRegex.test(question) && answer) {
          userPicks.pickem[question] = answer;
        }
      }
    });
    weeklyPicksCache[memberId] = userPicks;
  });

  return weeklyPicksCache;
}


/**
 * Fetches live scoreboard data to determine which games have ALREADY started.
 * @returns {Array<string>} An array of matchup short names, e.g., ["ARI @ LAR", "BUF @ MIA"].
 */
function getInvalidPickMatchups() {
  try {
    const response = espnFetch(SCOREBOARD); // Your global SCOREBOARD constant
    const data = JSON.parse(response.getContentText()).events;
    const pastGames = [];
    
    for (const event of data) {
      // "pre" means the game has not started.
      if (event.status.type.state != "pre") {
        pastGames.push(event.shortName);
      }
    }
    return pastGames;
  } catch (err) {
    Logger.log(`⚠️ Could not fetch scoreboard data: ${err.stack}`);
    return []; // Return an empty array on failure
  }
}

/**
 * Populates a contest sheet (Survivor or Eliminator) with the latest picks for a given week.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {Object} parsedPicks The clean "picks cache" object from the parser.
 * @param {Object} memberData The complete 'members' object.
 * @param {Object} config The main 'configuration' object.
 * @param {Object} gamePlan The game plan for the week, containing spread data.
 * @param {number} week The week number to populate.
 * @param {string} contestType The type of sheet to populate: 'survivor' or 'eliminator'.
 */
function populateSurvElimSheet(ss, parsedPicks, memberData, config, gamePlan, week, contestType) {
  const sheetName = contestType.toUpperCase();
  try {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = survElimSheet(ss, config, memberData, contestType);

    const isAts = gamePlan[`${contestType}Ats`];
    const weekColumn = parseInt(week) + 4; 
    sheet.setColumnWidth(weekColumn, isAts ? 68 : 42);

    // 1. Map Names to IDs
    const nameToIdObject = {};
    for (const id in memberData.members) {
      const name = memberData.members[id].name;
      if (name) nameToIdObject[name.toString().trim().toLowerCase()] = id;
    }

    // 2. Map IDs to Sheet Rows
    const memberIdToRowMap = {};
    const namesOnSheet = ss.getRangeByName(`${sheetName}_NAMES`).getValues().flat();
    namesOnSheet.forEach((name, index) => {
      const memberId = nameToIdObject[name.toString().trim().toLowerCase()];
      if (memberId) memberIdToRowMap[memberId] = index;
    });

    // 3. Define Logic Variables
    const pickKey = (contestType === 'survivor') ? 'sP' : 'eP';
    const weekIdx = parseInt(week) - 1;
    const dataRange = ss.getRangeByName(`${sheetName}_PICKS`);
    const totalRows = dataRange.getNumRows();
    let writeArray = Array(totalRows).fill(['']);

    // 4. Process Picks
    for (const memberId in parsedPicks) {
      const rowIndex = memberIdToRowMap[memberId];
      if (rowIndex === undefined) continue;

      const pickValue = parsedPicks[memberId]?.[contestType];
      writeArray[rowIndex] = [pickValue || ''];
      
      // Update the memberData Object JSON
      if (memberData.members[memberId] && pickValue) {
        let member = memberData.members[memberId];
        if (!member[pickKey]) member[pickKey] = [];
        while (member[pickKey].length < weekIdx) { member[pickKey].push(null); }
        member[pickKey][weekIdx] = pickValue;
      }
    }

    // 5. Write to Sheet
    sheet.getRange(2, weekColumn, totalRows, 1).setValues(writeArray);

    // 6. Save the Object
    saveProperties('members', memberData);

    const text = `Successfully populated ${sheetName} sheet and updated member records for Week ${week}.`;
    Logger.log(`✅ ${text}`);
    ss.toast(text,`✅ ${sheetName} IMPORTED`);
    return true;
  } catch (err) {
    const text = `❗ Failed to populate ${sheetName} sheet and member records for Week ${week}.`;
    Logger.log(text + '| ERROR: ' + err.stack);
    ss.toast(text,`${sheetName} PICK IMPORT FAILURE`);
  }
}


/**
 * Updates an existing Survivor/Eliminator sheet with the latest
 * data for all members after an outcome has been processed.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {config} is the configuration object from document properties
 * @param {memberData} is the 'members' document property object
 * @param {string} contestType The type of sheet to update: 'survivor' or 'eliminator'.
 */
function updateSurvElimSheet(ss, config, memberData, contestType) {
  contestType = contestType.toLowerCase(); // 'survivor' or 'eliminator'
  
  config = config || (JSON.parse(PropertiesService.getDocumentProperties().getProperty('configuration')) || {});
  memberData = memberData || (JSON.parse(PropertiesService.getDocumentProperties().getProperty('members')) || {});

  const sheetName = contestType.toUpperCase();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    const text = `❗ '${sheetName}' sheet not found, creating it now...`;
    const ui = SpreadsheetApp.getUi();
    ui.alert(`⭐ CREATING ${contestType.toUpperCase()} SHEET`, text, ui.ButtonSet.OK);
    sheet = survElimSheet(ss, null, null, contestType);
  }

  const namesRange = ss.getRangeByName(`${sheetName}_NAMES`);
  const livesRange = ss.getRangeByName(`${sheetName}_LIVES`);
  const revivesRange = ss.getRangeByName(`${sheetName}_REVIVES`);
  const eliminatedRange = ss.getRangeByName(`${sheetName}_ELIMINATED`);
  const picksRange = ss.getRangeByName(`${sheetName}_PICKS`);
  
  if (!namesRange || !picksRange) {
    Logger.log(`⚠️ Required named ranges for '${sheetName}' not found.`);
    return;
  }

  const memberNamesOnSheet = namesRange.getValues().flat();
  const picksData = picksRange.getValues(); 
  
  // Prepare Arrays for Batch Writing
  const newLivesData = [];
  const newRevivesData = [];
  const newEliminatedData = [];
  const newEliminatedBackgrounds = [];
  const newNameBackgrounds = [];
  const newNameFontLines = [];
  const newPickBackgrounds = []; 
  const newPickFontLines = [];
  const newPicksSummaryRow = Array(picksData[0].length).fill("");

  let poolLivesRemaining = 0;
  let totalRevives = 0;
  let membersRemaining = 0;

  // FIX: Get the correct lives count from config using the full key (e.g. survivorLives)
  const totalLivesConfig = parseInt(config[contestType + 'Lives'], 10) || 1;

  memberNamesOnSheet.forEach((name, rowIndex) => {
    const member = Object.values(memberData.members).find(m => m.name && m.name.toLowerCase() === name.toLowerCase());
    
    if (member) {
      const prefix = contestType.substring(0,1).toLowerCase(); // 's' or 'e'
      const livesArray = member[prefix + 'L'] || [];
      const currentLives = livesArray.length > 0 ? livesArray[livesArray.length - 1] : totalLivesConfig;

      // 1. Lives Dots Calculation
      const livesDots = totalLivesConfig > 1 
        ? '🟢'.repeat(currentLives) + '⚫'.repeat(Math.max(0, totalLivesConfig - currentLives)) 
        : (currentLives > 0 ? '🟢' : '❌');
      newLivesData.push([livesDots]);
      poolLivesRemaining += currentLives;

      // 2. Revives & Elimination Status
      const revivesEntry = member[prefix + 'R'];
      let memberRevivesCount = 0;

      if (Array.isArray(revivesEntry)) {
        // Sum the array, treating any null/undefined slots as 0
        memberRevivesCount = revivesEntry.reduce((sum, val) => sum + (parseInt(val) || 0), 0);
      } else {
        // Handle legacy integer or empty state
        memberRevivesCount = parseInt(revivesEntry) || 0;
      }

      newRevivesData.push([memberRevivesCount]);
      totalRevives += memberRevivesCount;
      
      const elimWeek = member[prefix + 'O'];
      newEliminatedData.push([elimWeek ? `OUT [WK${elimWeek}]` : 'IN']);
      newEliminatedBackgrounds.push([elimWeek ? '#ffccd6' : '#c7fcc7']);

      // 3. Name Formatting
      if (currentLives === 0) {
        newNameBackgrounds.push(['#ffccd6']); 
        newNameFontLines.push(['line-through']);
      } else {
        newNameBackgrounds.push(['#c7fcc7']); 
        newNameFontLines.push(['none']);
        membersRemaining++;
      }

      // 4. Pick Grid Formatting
      const evals = member[prefix + 'E'] || [];
      const rowPickColors = [];
      const rowPickFonts = [];

      for (let colIndex = 0; colIndex < picksData[0].length; colIndex++) {
        const pickOnSheet = picksData[rowIndex][colIndex];
        const isCorrect = evals[colIndex];
        
        if (!pickOnSheet || pickOnSheet.toString().trim() === "") {
          rowPickColors.push(null); 
          rowPickFonts.push('none');
        } else if (isCorrect === 1 || isCorrect === true) {
          if (newPicksSummaryRow[colIndex]) {
            newPicksSummaryRow[colIndex]++;
          } else {
            newPicksSummaryRow[colIndex] = 1;
          }
          rowPickColors.push('#c7fcc7'); // Green
          rowPickFonts.push('none');
        } else if (isCorrect === 0 || isCorrect === false) {
          rowPickColors.push('#ffccd6'); // Red
          rowPickFonts.push('line-through'); 
        } else {
          // Pending
          rowPickColors.push('#fffdd4'); // Yellow
          rowPickFonts.push('none');
        }
      }
      newPickBackgrounds.push(rowPickColors);
      newPickFontLines.push(rowPickFonts);

    } else {
      newLivesData.push(['']);
      newRevivesData.push(['']);
      newEliminatedData.push(['']);
      newNameBackgrounds.push([null]);
      newNameFontLines.push(['none']);
      newPickBackgrounds.push(Array(picksData[0].length).fill(null));
      newPickFontLines.push(Array(picksData[0].length).fill('none'));
    }
  });

  // EXECUTE BATCH UPDATES
  livesRange.setValues(newLivesData);
  revivesRange.setValues(newRevivesData);
  eliminatedRange.setValues(newEliminatedData);
  eliminatedRange.setBackgrounds(newEliminatedBackgrounds);
  namesRange.setBackgrounds(newNameBackgrounds);
  namesRange.setFontLines(newNameFontLines);
  picksRange.setBackgrounds(newPickBackgrounds);
  picksRange.setFontLines(newPickFontLines);

  // Update Totals
  sheet.getRange(livesRange.getLastRow() + 1, livesRange.getColumn()).setValue(poolLivesRemaining);
  sheet.getRange(revivesRange.getLastRow() + 1, revivesRange.getColumn()).setValue(totalRevives);
  sheet.getRange(eliminatedRange.getLastRow() + 1, eliminatedRange.getColumn()).setValue(membersRemaining);
  sheet.getRange(picksRange.getLastRow() + 1, picksRange.getColumn(), 1, picksRange.getNumColumns()).setValues([newPicksSummaryRow]);

  Logger.log(`✅ Refreshed '${sheetName}' visuals. Dots based on ${totalLivesConfig} max lives.`);
}

/**
 * FUNCTIONS FOR TRIGGER AND UPDATE SURVIVOR/ELIMINATOR STATUS
 */


/**
 * [RUN ONCE] Creates the installable onEdit trigger for the spreadsheet.
 * The user should be instructed to run this function once from the script editor
 * to enable automatic score and status processing.
 */
function createOnEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerFunctionName = 'onEditTrigger';

  // First, delete any existing triggers with the same function name to prevent duplicates.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === triggerFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create the new trigger to call our gatekeeper function.
  ScriptApp.newTrigger(triggerFunctionName)
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  SpreadsheetApp.getUi().alert('Success!', 'The automatic score processing trigger has been installed. Statuses will now update automatically when game outcomes are entered.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Finds and deletes the specific installable onEdit trigger used for
 * automatic score processing ('onEditTrigger'). This provides a clean way
 * for an admin to disable the feature.
 */
function deleteOnEditTrigger() {
  let triggerDeleted = false;

  // Get all triggers for the current project.
  const allTriggers = ScriptApp.getProjectTriggers();

  // Loop through the triggers to find the specific one we want to delete.
  for (const trigger of allTriggers) {
    // Identify trigger by the name of the function it is set to call.
    if (trigger.getHandlerFunction() === 'onEditTrigger') {
      
      // If found, delete it.
      ScriptApp.deleteTrigger(trigger);
      triggerDeleted = true;
      
      // Break the loop; assume there's only one.
      break; 
    }
  }

  // Provide clear feedback to the user.
  if (triggerDeleted) {
    SpreadApp.getUi().alert('Success', 'The automatic score processing trigger has been successfully removed.', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`❌ Automatic onEdit trigger was successfully deleted.`);
  } else {
    SpreadApp.getUi().alert('Info', 'No automatic score processing trigger was found to delete.', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log(`⭕ No onEdit trigger was found to delete.`);
  }
}

/**
 * This function is called by the installable onEdit trigger.
 * It efficiently checks if an edit was made to a relevant "outcome" cell on either
 * the main OUTCOMES sheet or a weekly sheet, then calls the main processing function.
 *
 * @param {Object} e The event object passed by the onEdit trigger.
 */
function onEditTrigger(e) {
  // 1. FAST GUARD: Is it a single cell?
  const range = e.range;
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  // 2. FAST GUARD: Get basic info without API calls where possible
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const editedRow = range.getRow();
  const editedCol = range.getColumn();
  
  let week = null;

  // --- Check 1: Main Outcomes Sheet ---
  if (sheetName === `${LEAGUE}_OUTCOMES`) {
    // Only care if editing row 4 or below, and if it's an outcome column
    if (editedRow > 3) {
      week = Math.ceil(editedCol / 2);
    }
  } 
  // --- Check 2: Weekly Sheet ---
  else if (sheetName.startsWith(weeklySheetPrefix)) {
    const weekRegex = new RegExp(`^${weeklySheetPrefix}(\\d{1,2})$`);
    const match = sheetName.match(weekRegex);
    if (match) {
      week = parseInt(match[1], 10);
      // We'll verify the specific range later inside the 'heavy lifter' to save time here
    }
  }

  if (week === null) return;

  // 3. SEMI-FAST GUARD: Now we do ONE property fetch
  const formsData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('forms') || "{}");
  if (!formsData[week]) return;

  // 4. CALL HEAVY LIFTER
  // Use a SpreadsheetApp toast to let the user know background work is happening
  e.source.toast(`Updating Week ${week} scores...`, "📊 Processing");
  
  try {
    // Pass the week and sheetName to your evaluator
    evalSurvElimStatus(week); 
    e.source.toast(`Week ${week} status updated!`, "✅ Success");
  } catch (err) {
    Logger.log(`Error: ${err.stack}`);
    e.source.toast("Update failed. Check logs.", "❌ Error");
  }
}


/**
 * Evaluates and updates Survivor and Eliminator statuses for a given week.
 *
 * @param {number} week The week number to process.
 * @param {string} sourceSheetName The name of the sheet that triggered the edit.
 */
function evalSurvElimStatus(week, sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docProps = PropertiesService.getDocumentProperties();
  
  // FIX 1: Load 'configuration' into config, NOT 'members'
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  const memberData = JSON.parse(docProps.getProperty('members')) || {};
  const formsData = JSON.parse(docProps.getProperty('forms')) || {};
  
  const weeklySheetName = `${weeklySheetPrefix}${week}`;
  const weeklySheet = ss.getSheetByName(weeklySheetName);

  // --- Step 1: Sync Outcome Data ---
  if (weeklySheet && sourceSheetName === weeklySheetName) {
    const weeklyOutcomesRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const masterOutcomesRange = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
    if (weeklyOutcomesRange && masterOutcomesRange) {
      masterOutcomesRange.setValues(weeklyOutcomesRange.getValues()[0].map(v => [v]));
    }
  }

  // --- Step 2: Gather Official Outcomes ---
  const winners = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`).getValues().flat();
  const margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`).getValues().flat();
  const gamePlan = formsData[week]?.gamePlan;

  if (!gamePlan) {
    Logger.log(`❌ Aborting: No gamePlan found for Week ${week}`);
    return;
  }

  const outcomeMap = new Map();
  gamePlan.games.forEach((game, index) => {
    const winnerRaw = winners[index];
    const winner = (winnerRaw && winnerRaw.toString().trim() !== "") ? winnerRaw.toString().trim() : null;
    const matchupKey = `${game.awayTeam} @ ${game.homeTeam}`;
    
    outcomeMap.set(matchupKey, {
      winner: winner,
      loser: (winner === 'TIE') ? 'TIE' : (winner === game.awayTeam ? game.homeTeam : game.awayTeam),
      margin: parseFloat(margins[index]) || 0,
      spread: game.spread
    });
  });

  Logger.log(`🏗 Map Built. Found ${Array.from(outcomeMap.values()).filter(o => o.winner).length} winners.`);

  // --- Step 3 & 4: Process Contests ---
  // FIX 2: Use Number() to ensure string "1" matches number 1
  if (config.survivorInclude && week >= Number(config.survivorStartWeek)) {
    Logger.log("🚀 Entering Survivor Processing...");
    processContest(ss, week, 'SURVIVOR', memberData, outcomeMap, config);
  }

  if (config.eliminatorInclude && week >= Number(config.eliminatorStartWeek)) {
    Logger.log("🚀 Entering Eliminator Processing...");
    processContest(ss, week, 'ELIMINATOR', memberData, outcomeMap, config);
  }
  
  // Save ALL updates
  saveProperties('members', memberData);
  saveProperties('configuration', config);

  // --- Step 6: Visual Refresh ---
  if (config.survivorInclude) updateSurvElimSheet(ss, config, memberData, 'survivor');
  if (config.eliminatorInclude) updateSurvElimSheet(ss, config, memberData, 'eliminator');

  ss.toast(`Week ${week} status updated!`, "✅ SUCCESS");
}

/**
 * [DEFINITIVE HELPER] A generic function to process either a Survivor or Eliminator contest.
 * This version uses a week-by-week array to track lives, enabling advanced history and revives.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {number} week The week number being processed.
 * @param {string} contestType The type of sheet to populate: 'SURVIVOR' or 'ELIMINATOR'.
 * @param {Object} memberData The complete 'members' object.
 * @param {Object} outcomeMap A Map of the final game outcomes for the week.
 * @param {Object} config The main 'configuration' object.
 * @returns {Object} The modified and updated memberData object.
 */
function processContest(ss, week, contestType, memberData, outcomeMap, config) {
  const sheet = ss.getSheetByName(contestType.toUpperCase());
  if (!sheet) return memberData;

  const names = ss.getRangeByName(`${contestType}_NAMES`).getValues().flat();
  const nameToIdMap = new Map();
  for (const id in memberData.members) {
    if (memberData.members[id]?.name) {
      nameToIdMap.set(memberData.members[id].name.toLowerCase(), id);
    }
  }

  const picksKey = contestType === 'SURVIVOR' ? 'sP' : 'eP';
  const evalKey = contestType === 'SURVIVOR' ? 'sE' : 'eE';
  const livesKey = contestType === 'SURVIVOR' ? 'sL' : 'eL';
  const outKey = contestType.toLowerCase().substring(0, 1) + 'O';
  
  const isAts = config[`${contestType.toLowerCase()}Ats`];
  const livesSetting = config[`${contestType.toLowerCase()}Lives`] || 1;
  const startWeek = config[`${contestType.toLowerCase()}StartWeek`] || 1;
  const picks = ss.getRangeByName(`${contestType}_PICKS`).getValues();

  let poolLives = 0, poolMembers = [], poolMembersEliminated = [];
  let pendingGamesDetected = false;

  names.forEach((name, rowIndex) => {
    const memberId = nameToIdMap.get(name.trim().toLowerCase());
    if (!memberId) return;

    const member = memberData.members[memberId];
    const pickAbbr = picks[rowIndex][week - 1]?.toString().match(/[A-Z]{2,3}/)?.[0];
    
    if (!pickAbbr) return;

    // --- Update Pick History ---
    if (!member[picksKey]) member[picksKey] = [];
    member[picksKey][week - 1] = pickAbbr;

    // --- Determine Outcome ---
    const gameKey = Array.from(outcomeMap.keys()).find(key => key.includes(pickAbbr));
    const outcome = outcomeMap.get(gameKey);

    if (!member[evalKey]) member[evalKey] = [];
    if (!member[livesKey]) member[livesKey] = [];

    let livesAtStartOfWeek = (week == startWeek) 
      ? parseInt(livesSetting, 10) 
      : (member[livesKey][week - 2] || 0);

    let isCorrect = null; // Default to null (Pending)

    if (!outcome || outcome.winner === null) {
      isCorrect = null; 
      pendingGamesDetected = true;
    } else {
      if (isAts) {
        isCorrect = calculateAtsResult(pickAbbr, outcome.winner, outcome.loser, outcome.margin, outcome.spread);
        if (contestType === 'ELIMINATOR') isCorrect = !isCorrect;
      } else {
        if (contestType === 'SURVIVOR') isCorrect = (outcome.winner === pickAbbr || outcome.winner === 'TIE');
        if (contestType === 'ELIMINATOR') isCorrect = (outcome.loser === pickAbbr);
      }
    }

    // Explicitly map True -> 1, False -> 0, and keep Null as Null
    member[evalKey][week - 1] = (isCorrect === null) ? null : (isCorrect ? 1 : 0);

    // --- Update Lives History ---
    let livesAtEndOfWeek = livesAtStartOfWeek;
    
    // This stays safe because (null === false) is false. 
    // Lives only drop if isCorrect is explicitly false.
    if (isCorrect === false && livesAtStartOfWeek > 0) {
      livesAtEndOfWeek--;
    }

    member[livesKey][week - 1] = livesAtEndOfWeek;

    // Handle elimination marker
    if (livesAtEndOfWeek === 0 && livesAtStartOfWeek > 0) {
      member[outKey] = week;
      Logger.log(`😵 ${member.name} eliminated from ${contestType} in Week ${week}`);
    } else if (livesAtEndOfWeek > 0) {
      delete member[outKey]; // Remove marker if they are revived or still in
    }

    // Counters for the completion check
    poolLives += livesAtEndOfWeek;
    if (livesAtEndOfWeek > 0) poolMembers.push(names[rowIndex]);
    if (livesAtStartOfWeek > 0 && livesAtEndOfWeek == 0) poolMembersEliminated.push(names[rowIndex]);
  });

  // --- Completion Check Logic ---
  // If games are still pending, we NEVER declare the pool finished.
  if (pendingGamesDetected) {
    Logger.log(`⏳ ${contestType}: Some games are still pending. Skipping completion check.`);
    config[`${contestType.toLowerCase()}Active`] = true;
  } else {
    let completionString;
    if (poolMembers.length > 1) {
      Logger.log(`🧮 ${contestType} lives remaining: ${poolLives}`);
      config[`${contestType.toLowerCase()}Active`] = true;
    } else {
      if (poolMembers.length == 1) {
        Logger.log(`🏆 ${contestType} COMPLETED! Winner: ${poolMembers[0]}`);
        completionString = `🏆 ${poolMembers[0]} is the sole remaining ${contestType} CHAMPION!`;
      } else {
        Logger.log(`🏆 ${contestType} COMPLETED! Last stand: ${poolMembersEliminated.join(', ')}`);
        completionString = `The pool ended with a multi-way elimination:\n🔹 ` + poolMembersEliminated.join(`\n🔹 `);
      }

      const activeKey = `${contestType.toLowerCase()}Active`;
      // If the key is undefined (first run) or true, show the alert
      if (config[activeKey] !== false) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(`✔️ ${contestType} COMPLETE!`, `${completionString}\n\nTo restart, update the Start Week in Config.`, ui.ButtonSet.OK);
      }
      config[`${contestType.toLowerCase()}Active`] = false;
    }
  }

  saveProperties('configuration', config);
  return memberData;
}

/**
 * Calculates if a pick was correct against the spread,
 * using only the winning team and the margin of victory.
 *
 * @param {string} pick The member's picked team abbreviation (e.g., "DAL").
 * @param {string} winner The actual winning team's abbreviation (e.g., "PHI").
 * @param {number} margin The positive margin of victory (e.g., 3).
 * @param {string} spread The original spread string from the game plan (e.g., "DAL -6.5").
 * @returns {boolean} True if the pick was correct against the spread.
 */
function calculateAtsResult(pick, winner, loser, margin, spread) {
  if (!spread || spread.toUpperCase() === 'PK' || spread.trim() === '0' || spread.toUpperCase() === 'PUSH' || spread.toUpperCase() === 'EVEN') {
    // If it's a "pick'em", the pick is correct if they picked the winner.
    return pick === winner;
  }

  try {
    const spreadMatch = spread.match(/([A-Z]{2,3})\s*([+-]?\d+\.?\d*)/);
    if (!spreadMatch) return false; // Invalid spread format

    const [, favoriteTeam, spreadValueStr] = spreadMatch;
    const spreadValue = parseFloat(spreadValueStr); // e.g., -6.5
    const underdogTeam = favoriteTeam == winner ? loser : winner;
    if (pick === underdogTeam) {
      return true;
    } else if (pick === favoriteTeam && margin > Math.abs(spreadValue)) {
      // They picked the FAVORITE. They win if the actual winner is the favorite AND the margin is greater than the spread.
      // e.g., Spread is -6.5, margin must be 7 or more.
      return true;
    } else {
      // There was a TIE or the favorite won by less than the spread value
      return false;
    }
  } catch (err) {
    Logger.log(`⚠️ Error calculating ATS result for spread "${spread}"| ERROR: ${err.stack}`);
    return false;
  }
}

/**
 * Wrapper function to call the sync process from the spreadsheet menu.
 */
function syncCurrentWeekResponses() {
  const ui = SpreadsheetApp.getUi();
  try {
    const week = fetchWeek(false, true);
    // Show a toast to indicate the process is starting.
    SpreadsheetApp.getActiveSpreadsheet().toast('Syncing responses for Week ' + week + '...', 'In Progress', 10);
    const result = syncFormResponses(week);
    ui.alert('Sync Complete', `${result.message}\nFound ${result.newMembers} new member(s).\nTotal Respondents: ${result.totalRespondents}.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to sync responses: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Orchestrates the entire process of fetching, de-duplicating,
 * and processing form responses to update the 'members' and 'forms' properties.
 */
function syncFormResponses(week) {
  // SETUP 
  Logger.log(`🟩 Getting things ready to import for week ${week}`);
  week = week || fetchWeek(null,true);
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration'));
  let memberData = JSON.parse(docProps.getProperty('members')) || { memberOrder: [], members: {}};
  let formsData = JSON.parse(docProps.getProperty('forms'));
  const responseSheet = getDatabaseSheet().getSheetByName(`WK${week}`);
  
  if (!responseSheet) {
    Logger.log(`⭕ No response sheet found for WK${week}.`);
    return { success: true, message: `⭕ No response sheet exists for Week ${week}.`, newMembers: 0, totalRespondents: 0 };
  }
  
  const data = responseSheet.getDataRange().getValues();

  if (data.length < 2) {
    Logger.log(`⭕ No responses found in sheet for WK${week}.`);
    return { success: true, message: `⭕ No responses to sync for Week ${week}.`, newMembers: 0, totalRespondents: 0 };
  }

  const headers = data.shift();
  const { nameCol, newNameCol } = findNameColumns(headers);
  const newUserAnswerRegex = /new user/i;
  
  // DE-DUPLICATE RESPONSES (Latest Submission Used) ---
  const latestSubmissions = {};
  data.forEach(row => {
    const name = (nameCol == -1 || newUserAnswerRegex.test(row[nameCol])) ? row[newNameCol] : row[nameCol];
    if (name && name.trim() !== '') {
      latestSubmissions[name.trim().toLowerCase()] = row;
    }
  });
  const finalResponseRows = Object.values(latestSubmissions);
  // PROCESS NEW MEMBERS
  const nameToIdMap = {};
  for (const id in memberData.members) {
    nameToIdMap[memberData.members[id].name.toLowerCase()] = id;
  }
  const newMemberIds = [];
  
  finalResponseRows.forEach(row => {
    const submitterChoice = row[nameCol];
    const newUserName = (nameCol == -1) ? row[newNameCol].trim() : (newNameCol > -1) ? row[newNameCol].trim() : '';
    
    if ((nameCol == -1 || newUserAnswerRegex.test(submitterChoice)) && newUserName) {
      const nameKey = newUserName.trim().toLowerCase();
      Logger.log(`🔄 Processing user named ${newUserName}, generating unique id...`)
      // If their name is NOT in official map, consider user a new member.
      if (!nameToIdMap[nameKey]) {
        const permanentId = generateUniqueId();
        
        memberData.memberOrder.push(permanentId);
        memberData.members[permanentId] = createNewMember(
          newUserName,
          false,
          config,
          week // The join week is the week of the form they submitted
        );
        
        nameToIdMap[nameKey] = permanentId;
        Logger.log(`🆗 User ${newUserName}, stored under ${nameKey} and given unique ID of ${permanentId}`)
        newMemberIds.push(permanentId);
      }
    }
  });

  // UPDATE RESPONDENT LIST IN 'forms' OBJECT ---
  const respondentIds = finalResponseRows.map(row => {
    const submitterChoice = row[nameCol];
    const name = (nameCol == -1) ? row[newNameCol] : newUserAnswerRegex.test(submitterChoice) ? ((newNameCol > -1) ? row[newNameCol] : '') : submitterChoice;
    return nameToIdMap[name.trim().toLowerCase()];
  }).filter(id => id);
  
  if (formsData[week]) {
    formsData[week].respondents = [...new Set(respondentIds)];
    if (newMemberIds.length > 0) {
      if (formsData[week].hasOwnProperty('newMembers')) {
        formsData[week].newMembers.push(...newMemberIds);
      } else {
        formsData[week].newMembers = newMemberIds;
      }
    }
    formsData[week].responseCount = formsData[week].respondents.length;
    formsData[week].lastResponseTime = new Date().toISOString();
    const allMemberIds = memberData.memberOrder;
    formsData[week].nonRespondents = allMemberIds.filter(id => !respondentIds.includes(id));
  }

  formsData[week].imported = false;

  // SAVE UPDATED DATA ---
  saveProperties('members', memberData);
  saveProperties('forms', formsData);

  // const newMembers = newMemberIds.map(id => memberData.members[id]?.name);
  return {
    success: true,
    message: `Sync complete for Week ${week}.`,
    newMembers: formsData[week].newMembers,
    totalRespondents: respondentIds.length
  };
}

/**
 * Finds the column indexes for name fields using regex.
 */
function findNameColumns(headers) {
  const mainNameRegex = /select.*name/i; 
  const newUserNameRegex = /enter.*name/i;
  let nameCol = -1, newNameCol = -1;

  headers.forEach((header, index) => {
    if (mainNameRegex.test(header)) nameCol = index;
    else if (newUserNameRegex.test(header)) newNameCol = index;
  });
  
  return { nameCol, newNameCol };
}

/** 
 * Function to eliminate the "New User" entry on the specified week's form
 */
function removeNewUserQuestion(week) {
  let nameQuestion, found = false;
  try {
    const id = fetchProperties('forms')[week].id;
    let form = FormApp.openById(id);
    let items = form.getItems();
    for (let a = 0; a < items.length; a++) {
      if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
        nameQuestion = items[a];
      }
    }

    let choices = nameQuestion.asListItem().getChoices();
    for (let a = 0; a < choices.length; a++) {
      if (choices[a].getValue() == 'New User') {
        choices.splice(a,1);
        found = true;
      }
    }
    if (found) {
      nameQuestion.asListItem().setChoices(choices);
      ss.toast('Removed the option of \"New User\" from the form.');
    } else {
      ss.toast('No \"New User\" option was present on the form.');
    }
  }
  catch (err) {
    ss.toast('Failed to remove the list item of \"New User\" from the form.');
  }
}

// ============================================================================================================================================
// CORE SHEETS
// ============================================================================================================================================

/**
 * Creates or updates the OUTCOMES sheet with a multi-row header, 
 * custom color scales for margins, and robust handling of empty playoff weeks.
 */
function outcomesSheet(ss) {
  ss = fetchSpreadsheet(ss);
  const sheetName = `${LEAGUE}_OUTCOMES`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // --- Start with a clean slate ---
  sheet.clear();
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearDataValidations().clearNote();
  sheet.getNamedRanges().forEach(namedRange => {
    if (namedRange.getName().startsWith(`${LEAGUE}_OUTCOMES`)) {
      namedRange.remove();
    }
  });
  sheet.setTabColor(dayColorsFilled[dayColorsFilled.length - 1] || null); // Your custom tab color

  const data = ss.getRangeByName(LEAGUE)?.getValues() || fetchSchedule(ss);
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  // --- 1. Build the New, 3-Row Header Structure ---
  const weekTypeHeaders = []; // For Row 2 (e.g., "Regular Season", "WildCard")
  const weekNumHeaders = [];  // For Row 3 (e.g., "Week 1", "Week 2")

  for (const a in weeks) {
    const weekName = WEEKNAME[weeks[a]] ? WEEKNAME[weeks[a]].name : "Regular Season";
    weekTypeHeaders.push(weekName,""); // Add name and a blank for the margin column
    weekNumHeaders.push(`Week ${weeks[a]}`,""); // Add week # and a blank for the margin column|
  }
  
  // --- 2. Sheet Resizing and Basic Formatting ---
  const headerRow1 = 1; // Main Title
  const headerRow2 = 2; // Week Type (e.g., WildCard)
  const headerRow3 = 3; // Week Number
  const dataStartRow = 4; // Data starts on row 4 now

  const totalCols = weekNumHeaders.length;
  const totalRows = dataStartRow + MAXGAMES -1;
  
  if (sheet.getMaxColumns() < totalCols) sheet.insertColumnsAfter(1, totalCols - 1);
  if (sheet.getMaxRows() < totalRows) sheet.insertRowsAfter(1, totalRows - 1);
  if (sheet.getMaxColumns() > totalCols) sheet.deleteColumns(totalCols+1,sheet.getMaxColumns()-totalCols);
  if (sheet.getMaxRows() > totalRows) sheet.deleteRows(totalRows+1,sheet.getMaxRows()-totalRows);
  
  sheet.getRange(headerRow3+1,1,totalRows-headerRow3,weekNumHeaders.length).setBackground('#dddddd');

  // --- 3. Apply Header and Base Styles ---
  // Main Title (Row 1)
  sheet.getRange(headerRow1, 1, 1, totalCols).mergeAcross().setValue(sheetName.replace(/_/g, ' '))
      .setFontWeight('bold').setFontSize(18).setFontFamily("Montserrat")
      .setBackground('#666666').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow1, 40);

  // Week Type Headers (Row 2)
  sheet.getRange(headerRow2, 1, 1, totalCols).setValues([weekTypeHeaders]).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setFontStyle('italic').setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow2, 25);
  
  // Week Number Headers (Row 3)
  sheet.getRange(headerRow3, 1, 1, totalCols).setValues([weekNumHeaders])
      .setBackground('#000000').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow3, 25);
  
  // Merge the header cells in pairs
  for (let i = 1; i <= totalCols; i += 2) {
      sheet.getRange(headerRow2, i, 1, 2).mergeAcross();
      sheet.getRange(headerRow3, i, 1, 2).mergeAcross();
  }

  const dataBodyRange = sheet.getRange(dataStartRow, 1, MAXGAMES, totalCols);
  dataBodyRange.setFontFamily("Montserrat").setFontSize(9).setVerticalAlignment('middle').setHorizontalAlignment('center');

  // --- 4. Populate Matchups, Set Ranges, Validation, and Formatting ---
  let allConditionalFormatRules = [];
  const marginValidationRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 50).build();
  let col = 1;
  for (const a in weeks) {
    const winnerCol = (col - 1) * 2 + 1;
    const marginCol = winnerCol + 1;
    sheet.setColumnWidth(winnerCol, 50);
    sheet.setColumnWidth(marginCol, 50);

    // [THE FIX] Filter for the week's matchups, gracefully handling empty weeks.
    const weekMatchups = data.filter(row => row[0] == weeks[a]);
    
    // If no matchups for this week (e.g., future playoff week), skip to the next loop iteration.
    if (weekMatchups.length === 0) {
      continue; 
    }

    const winnerRange = sheet.getRange(dataStartRow, winnerCol, weekMatchups.length);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${weeks[a]}`, winnerRange);
    const marginRange = sheet.getRange(dataStartRow, marginCol, weekMatchups.length);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${weeks[a]}_MARGIN`, marginRange);
    marginRange.setDataValidation(marginValidationRule);

    weekMatchups.forEach((game, index) => {
      const rowIndex = dataStartRow + index;
      const winnerCell = sheet.getRange(rowIndex, winnerCol);
      const marginCell = sheet.getRange(rowIndex, marginCol);
      const awayTeam = game[6];
      const homeTeam = game[7];
      const dayIndex = game[2] + 4; // Numeric day used for gradient application (-4 is Wednesday, 1 is Monday);
      
      winnerCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([awayTeam, homeTeam, 'TIE'], true).setAllowInvalid(false).build());
      marginCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());

      // --- [THE NEW COLOR LOGIC] ---
      // Set the base background color for BOTH cells based on the day of the week.
      winnerCell.setBackground(dayColors[dayIndex]);
      marginCell.setBackground(dayColors[dayIndex]);

      // Create a conditional format rule for the winning team's background color.
      // This will override the base color only when a winner is selected.
      const homeWinRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(homeTeam)
        .setBold(true)
        .setBackground(dayColorsFilled[dayIndex] || null)
        .setRanges([winnerCell])
        .build();
      allConditionalFormatRules.push(homeWinRule);
      
      const awayWinRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(awayTeam)
        .setBold(false) // Home team is bold, away is not
        .setBackground(dayColorsFilled[dayIndex] || null)
        .setRanges([winnerCell])
        .build();
      allConditionalFormatRules.push(awayWinRule);

      // [BONUS] Custom color scale for the margin, using the day's colors.
      // It will scale from the lighter day color to the darker filled day color.
      const marginColorScaleRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue(dayColors[dayIndex] || '#e0e0e0', SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMaxpointWithValue(dayColorsFilled[dayIndex] || '#b0b0b0', SpreadsheetApp.InterpolationType.NUMBER, '10')
        .setRanges([marginCell])
        .build();
      allConditionalFormatRules.push(marginColorScaleRule);
    });
    col++;
  }
  
  // Add a single rule for TIEs for all columns
  const tieRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TIE')
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([sheet.getRange(dataStartRow,1,sheet.getMaxRows()-dataStartRow,sheet.getLastColumn())])
    .build();
  allConditionalFormatRules.unshift(tieRule);
  // Add a single rule for TIEs for all columns
  const zeroRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([sheet.getRange(dataStartRow,1,sheet.getMaxRows()-dataStartRow,sheet.getLastColumn())])
    .build();
  allConditionalFormatRules.unshift(zeroRule);
  sheet.setConditionalFormatRules(allConditionalFormatRules);

  Logger.log(`✅ Completed setting up ${LEAGUE} OUTCOMES sheet`);
}


/**
 * UPDATE OUTCOMES SHEET NAMED RANGES BASED ON FORM WITH PROMPT
 * 
 * Allows user input as a utility to fix broken or outdated data ranges
 * 
 */
function outcomesSheetUpdatePrompt() {
  const ss = fetchSpreadsheet();
  const ui = fetchUi();
  
  let outcomesSheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);
  if (!outcomesSheet) {
    Logger.log(`⚠️ No ${LEAGUE}_OUTCOMES sheet found...`);
    ss.toast(`No ${LEAGUE}_OUTCOMES sheet found...`,`⚠️ ERROR`);
    if (ui.prompt(`❗ NO ${LEAGUE}_OUTCOMES SHEET FOUND`,`The sheet needed for this function to operate on was not found. Create the ${LEAGUE}_OUTCOMES sheet now?`, ui.ButtonSet.OK_CANCEL) == 'OK') {
      outcomesSheet();
    } else {
      Logger.log(`🚫 Outcomes sheet update function canceled by user. No OUTCOMES sheet present.`)
      return null;
    }
  }
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  let weekPrompt = ui.prompt(`📅 SUBMIT WEEK NUMBER`,`Enter a week number for which to rebuild ${LEAGUE}_OUTCOMES sheet column validation.\n\n⚠️ CAUTION: This function should be unnecessary if you've imported the weeks' picks. Be advised that this function will erase outcomes for the week provided in the ${LEAGUE}_OUTCOMES sheet.\n\nAvailable entries to select from: ${weeks.join(', ')}`,ui.ButtonSet.OK_CANCEL);
  let value, emojiString, invalid = true;
  while (weekPrompt.getSelectedButton() == 'OK' && invalid) {
    value = parseInt(weekPrompt.getResponseText());
    if (weeks.indexOf(value) >= 0) {
      emojiString = value > 10 ? numberMap[Math.floor(value/10)] + numberMap[value - Math.floor(value/10)*10] : numberMap[value];
      ss.toast(`Submitted week value of ${value} to rebuild outcomes for.`,emojiString);
      Logger.log(`${emojiString} You submitted a valid entry of ${value}`);
      invalid = false;
    } else {
      Logger.log(`❗ You submitted an invalid entry of ${weekPrompt.getResponseText()}`);
      weekPrompt = ui.prompt(`📅 RE-SUBMIT WEEK NUMBER`,`Please try again. Available entries to select from: ${weeks.join(', ')}`,ui.ButtonSet.OK_CANCEL);
    }
  }
  if (invalid) {
    ss.toast(`Canceled OUTCOMES sheet update by user (selected the "Cancel" button)`,`🚫 CANCELED`);
    Logger.log(`🚫 Canceled OUTCOMES sheet update by user (selected the "Cancel" button)`);
    return null;
  }

  let confirm = ui.alert(emojiString,`You submitted a value of ${value}, proceed with rebuild of this week's formatting within the ${LEAGUE}_OUTCOMES sheet?`,ui.ButtonSet.YES_NO);
  if (confirm == 'YES') {
    outcomesSheetUpdate(ss,value,null,null);
    ss.toast(`Successfully updated ${LEAGUE}_OUTCOMES sheet with data validation and coloration for week ${value}.`,`🎉 SUCCESS!`);
    Logger.log(`🎉 Successfully updated ${LEAGUE}_OUTCOMES sheet with data validation and coloration for week ${value}.`);
  } else {
    ss.toast(`Update of week ${value} OUTCOMES sheet columns canceled by user.`,`🚫 CANCELED`);
    Logger.log(`🚫 Canceled OUTCOMES sheet update by user. Deferred confirmation of script.`);
  }
}

/** 
 * UPDATE OUTCOMES - Updates the data validation, color scheme, and matchups for a specific week on the winners sheet
 * 
 * Modified to ensure that the matchups displayed are parallel to those from the weekly sheet -- only called by the "weeklySheet" function
 * 
*/
function outcomesSheetUpdate(ss,week,config,gamePlan) {
  ss = ss || fetchSpreadsheet(ss);
  week = week || fetchWeek();
  let sheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);
  let matchups = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
  let margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);
  if (!sheet || !matchups) {
    const missingOutcomes = `⚠️ Outcomes sheet error or not present, creating now`;
    Logger.log(missingOutcomes);
    ss.toast('OUTCOMES SHEET ISSUE',missingOutcomes);
    sheet = outcomesSheet(ss);
    matchups = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
    margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);
  }
  const startRow = matchups.getRow(); // First row of matchups on 

  let docProps;
  if (!config || !gamePlan) {
    docProps = PropertiesService.getDocumentProperties();
  }
  gamePlan = gamePlan || JSON.parse(docProps.getProperty('forms'))[week].gamePlan || {};
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  
  const contests = gamePlan.games;
  
  // Clears data validation and notes
  
  matchups.clearDataValidations().clearNote();
  margins.clearDataValidations().clearNote();
  
  let existingRules = sheet.getConditionalFormatRules();
  let rulesToKeep = [];
  let newRules = [];
  for (let a = 0; a < existingRules.length; a++) {
    let ranges = existingRules[a].getRanges();
    for (let b = 0; b < ranges.length; b++) {
      if (ranges[b].getColumn() != matchups.getColumn() && ranges[b].getColumn() != margins.getColumn()) {
        rulesToKeep.push(existingRules[a]);
      }
    }
  }
  sheet.clearConditionalFormatRules();
  
  sheet.getRange(startRow,matchups.getColumn(),sheet.getMaxRows()-startRow+1,1).setBackground('#dddddd');
  sheet.getRange(startRow,margins.getColumn(),sheet.getMaxRows()-startRow+1,1).setBackground('#dddddd');
  let start = startRow;
  let end = start+1;
  let teams = []; // Array to cross-reference existing values when re-writing
  
  for (let a = 0; a < contests.length; a++) {
    teams.push(contests[a].awayTeam);
    teams.push(contests[a].homeTeam);
    sheet.getRange(a+startRow,matchups.getColumn()).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([contests[a].awayTeam,contests[a].homeTeam,'TIE'], true).build());
    // Color Coding Days
    if (contests[a].dayName != contests[a+1]?.dayName) {
      // Matchup column color (static) and conditional formatting
      matchupCell = sheet.getRange(start,matchups.getColumn(),end-start,1)
      matchupCell.setBackground(dayColorsObj[contests[a].dayName] || '#e0e0e0');
      let homeWin = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=iferror(match(indirect("R[0]C[0]",false),indirect("${LEAGUE}_HOME_${week}"),0)>=0,false)`)
        .setBackground(dayColorsFilledObj[contests[a].dayName] || '#b0b0b0')
        .setBold(true)
        .setRanges([matchupCell])
        .build();
      newRules.push(homeWin);
      let awayWin = SpreadsheetApp.newConditionalFormatRule()
        .whenCellNotEmpty()
        .setBackground(dayColorsFilledObj[contests[a].dayName] || '#b0b0b0')
        .setRanges([matchupCell])
        .build();
      newRules.push(awayWin);

      // Margin column color (static) and conditional formatting
      marginCell = sheet.getRange(start,margins.getColumn(),end-start,1);
      marginCell.setBackground(dayColorsObj[contests[a].dayName] || '#e0e0e0');
      const marginColorScaleRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue(dayColorsObj[contests[a].dayName] || '#e0e0e0', SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMaxpointWithValue(dayColorsFilledObj[contests[a].dayName] || '#b0b0b0', SpreadsheetApp.InterpolationType.NUMBER, '10')
        .setRanges([marginCell])
        .build();
      newRules.push(marginColorScaleRule);

      start = end;

    }
    end++;
  }

  sheet.getRange(startRow,margins.getColumn(),contests.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());
  
  let allRules = rulesToKeep.concat(newRules);
  //clear all rules first and then add again
  
  sheet.setConditionalFormatRules(allRules);

  if (config.pickemsInclude) {
    // This function subcomponent runs when there are pick 'ems present and ties the response cell in Outcomes sheet to the response within the weekly sheet in question. It also prevents overwriting the values that may exist in the outcomes sheet, if present.
    let weeklySheetName = (weeklySheetPrefix + week);
    
    let sourceSheet = ss.getSheetByName(weeklySheetName);
    const targetSheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);

    const sourceMatchupRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const sourceMarginRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`);
    const targetMatchupRange = targetSheet.getRange(startRow,matchups.getColumn(),contests.length,1);
    const targetMarginRange = targetSheet.getRange(startRow,margins.getColumn(),contests.length,1);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${week}`,targetMatchupRange);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${week}_MARGIN`,targetMarginRange);

    let matchupCol = targetSheet.getRange(startRow,matchups.getColumn(),targetSheet.getMaxRows()-startRow+1,1);
    let dataMatchups = matchupCol.getValues().flat();
    let marginCol = targetSheet.getRange(startRow,margins.getColumn(),targetSheet.getMaxRows()-startRow+1,1);
    let dataMargins = marginCol.getValues().flat();
    let regexMatchups = new RegExp(/^[A-Z]{2,3}/);
    let regexMargins = new RegExp(/^[0-9]{1,2}/);
    let reWriteMatchups = [], reWriteMargins = [];
    // Data retention (if present)
    let map = [];
    for (let a = 0; a <= dataMatchups.length; a++) {
      if (regexMatchups.test(dataMatchups[a])) {
        reWriteMatchups.push(dataMatchups[a]);
        reWriteMargins.push(dataMargins[a]);
        map.push('');
      }
    }

    for (let a = 0; a < reWriteMatchups.length; a++) {
      if (teams.indexOf(reWriteMatchups[a])) {
        map[a] = Math.floor(teams.indexOf(reWriteMatchups[a])/2);
      }
    }
    matchupCol.clearContent();
    marginCol.clearContent();

    for (let e = map.length - 1; e >= 0; e--) {
      targetSheet.getRange(targetMatchupRange.getRow()+map[e],targetMatchupRange.getColumn()).setValue(reWriteMatchups[e]);
      targetSheet.getRange(targetMarginRange.getRow()+map[e],targetMarginRange.getColumn()).setValue(reWriteMargins[e]);
    }

    for (let a = 1; a <= sourceMatchupRange.getNumColumns(); a++) {
      if (!regexMatchups.test(dataMatchups[a-1])) {
        targetSheet.getRange(targetMatchupRange.getRow()+(a-1),targetMatchupRange.getColumn()).setFormula(
          '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(sourceMatchupRange.getRow(),sourceMatchupRange.getColumn()+(a-1)).getA1Notation()
        );
      } else {
        Logger.log(`Found matching matchup outcome value of ${dataMatchups[a-1]} on outcomes sheet in row ${(a + 2)}; avoiding re-writing formula for this cell`);
      }
    }
    for (let a = 1; a <= sourceMarginRange.getNumColumns(); a++) {
      if (!regexMargins.test(dataMargins[a-1])) {
        targetSheet.getRange(targetMarginRange.getRow()+(a-1),targetMarginRange.getColumn()).setFormula(
          '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(sourceMarginRange.getRow(),sourceMarginRange.getColumn()+(a-1)).getA1Notation()
        );
      } else {
        Logger.log(`Found matching outcome margin value of ${dataMargins[a-1]} on outcomes sheet in row ${(a + 2)}; avoiding re-writing formula for this cell`);
      }
    }
  }
}


/** 
 * TOTAL Sheet Creation / Adjustment
*/
function totSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);
  
  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;
  
  let sheetName = 'TOTAL';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  
  let rows = totalMembers+2;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }

  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(2,1).setValue('AVERAGES');

  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,30);
    sheet.getRange(2,a+3).setFormula('=iferror(arrayformula(countif(filter('+LEAGUE+'_PICKS_'+(weeks[a])+',NAMES_'+(weeks[a])+'=$A2)='+LEAGUE+'_PICKEM_OUTCOMES_'+(weeks[a])+',true)),)');
  }
  
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6');
  
  sheet.getRange(2,2,totalMembers+1,weeks.length+1).setNumberFormat('#.#');

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  // SET OVERALL NAMES Range
  let rangeOverallTotNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_NAMES',rangeOverallTotNames); 
  sheet.clearConditionalFormatRules(); 
  // OVERALL TOTAL GRADIENT RULE
  let rangeOverallTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL',rangeOverallTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL"))') // Min value of all correct picks
    .setRanges([rangeOverallTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY',range);
  let formatRuleOverallHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "15")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallHigh);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
  overallMainFormulas(weeks,sheet,totalMembers,'TOT',true);
  
  return sheet;  
}

// RNK Sheet Creation / Adjustment
function rnkSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);
  
  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'RNK';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  sheet.setTabColor(generalTabColor);

  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('RANKS');
  sheet.getRange(1,2).setValue('AVERAGE');

  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,48);
  }
    
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL RANK NAMES Range
  let rangeOverallTotRnkNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_RNK_NAMES',rangeOverallTotRnkNames);  
  sheet.clearConditionalFormatRules(); 
  // RANKS TOTAL GRADIENT RULE
  let rangeOverallRankTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_RANK',rangeOverallRankTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRankTot])
    .build();
  // RANKS SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY_RANK',range);
  let formatRuleOverallWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#00E1FF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallWinner);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
  overallMainFormulas(weeks,sheet,totalMembers,'RNK',false);
  
  return sheet;  
}

// PCT Sheet Creation / Adjustment
function pctSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);

  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'PCT';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  
  let rows = totalMembers+2; // 2 additional rows
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('PERCENTAGES');
  sheet.getRange(1,2).setValue('AVERAGE');
  sheet.getRange(rows,1).setValue('AVERAGES');
  
  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,48);
  }
  
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6'); 

  sheet.getRange(2,2,totalMembers+1,1).setNumberFormat("##.#%");  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL PCT NAMES Range
  let rangeOverallTotPctNames = sheet.getRange('R2C1:R'+(rows-1)+'C1');
  ss.setNamedRange('TOT_OVERALL_PCT_NAMES',rangeOverallTotPctNames);
  sheet.clearConditionalFormatRules();
  // PCT TOTAL GRADIENT RULE
  let rangeOverallTotPct = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_OVERALL_PCT',rangeOverallTotPct);
  rangeOverallTotPct = sheet.getRange('R2C2:R'+rows+'C2');
  let formatRuleOverallPctTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_PCT"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_PCT"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_PCT"))') // Min value of all correct picks  
    .setRanges([rangeOverallTotPct])
    .build();  
  // PCT SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY_PCT',range);
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2)); 
  let formatRuleOverallPctHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverallPct = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallPctHigh);
  formatRules.push(formatRuleOverallPct);
  formatRules.push(formatRuleOverallPctTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
  overallMainFormulas(weeks,sheet,totalMembers,'PCT',true);

  return sheet;  
}

// MNF Sheet Creation / Adjustment
function mnfSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);

  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'MNF';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  Logger.log(`📡 Checking for Monday games, if any`);
  let data = ss.getRangeByName(LEAGUE).getValues();
  let text = '0';
  let result = text.repeat(weeks.length);
  let mondayNightGames = Array.from(result);
  for (let a = 0; a < data.length; a++) {
    if ( data[a][2] == 1 && data[a][3] >= 17) {
      mondayNightGames[(data[a][0]-1)]++;
    }
  }
  let rows = totalMembers + 2; // AustinOrphan's suggestion!
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(rows,1).setValue('AVERAGES');

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);

  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6'); 
  
  let headers = [];
  for ( let a = 0; a < weeks.length; a++ ) {
    if (mondayNightGames[a] == 2) {
      range = sheet.getRange(1,a+3);
      range.setNote('Two MNF Games')
        .setFontWeight('bold')
        .setBackground('#555555');
    } else if (mondayNightGames[a] == 3) {
      range = sheet.getRange(1,a+3);
      range.setNote('Three MNF Games')
        .setFontWeight('bold')
        .setBackground('#999999');
    } else if (mondayNightGames[a] == 4) {
      range = sheet.getRange(1,a+3);
      range.setNote('Four MNF Games')
        .setFontWeight('bold')
        .setBackground('#CCCCCC');
    } else if (mondayNightGames[a] >= 4) {
      range = sheet.getRange(1,a+3);
      range.setNote(mondayNightGames[a] + ' MNF Games')
        .setFontWeight('bold')
        .setFontColor('black')
        .setBackground('#FFFFFF');
    }
    sheet.setColumnWidth(a+3,30);
    headers.push(weeks[a]);
  }
  sheet.getRange(1,3,1,weeks.length).setValues([headers]);

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  sheet.clearConditionalFormatRules(); 

  // SET MNF NAMES Range
  let rangeMnfNames = sheet.getRange(`R2C1:R${rows-1}C1`);
  ss.setNamedRange('MNF_NAMES',rangeMnfNames); 
  // MNF TOTAL GRADIENT RULE
  let rangeMnfTot = sheet.getRange(`R2C2:R${rows-1}C2`);
  ss.setNamedRange('MNF',rangeMnfTot);
  let formatRuleMnfTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("MNF"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("MNF"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("MNF"))') // Min value of all correct picks
    .setRanges([rangeMnfTot])
    .build();
  // MNF AVERAGES GRADIENT RULE
  let rangeMnfAvg = sheet.getRange(`R${rows}C2:R${rows}C${weeks.length+2}`);
  let formatRuleMnfAvg = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([rangeMnfAvg])
    .build();
  // MNF SHEET GRADIENT RULE
  range = sheet.getRange(`R2C3:R${rows-1}C${weeks.length+2}`);
  ss.setNamedRange('MNF_WEEKLY',range);
  let formatRuleTwoCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground('#9CFFC4')
    .setFontColor('#9CFFC4')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOneCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#C9FFDF')
    .setFontColor('#C9FFDF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=or(and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R[0]C[0]\",false)=0),and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"WEEK\")>=indirect(\"R1C[0]\",false)))')
    .setBackground('#FFC4CA')
    .setFontColor('#FFC4CA')
    .setBold(true)
    .setRanges([range])
    .build();    
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleTwoCorrect);
  formatRules.push(formatRuleOneCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleMnfTot);
  formatRules.push(formatRuleMnfAvg);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false);
  overallMainFormulas(weeks,sheet,totalMembers,'MNF',true);

  return sheet;  
}


/**
 * Sheet creation tool for the survivor and eliminator sheets
 * 
 */
function survElimSheet(ss,config,memberData,sheetType) {
  ss = ss || fetchSpreadsheet(ss);
  let docProps;
  if (!config || !memberData) docProps = PropertiesService.getDocumentProperties();

  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  
  sheetType = sheetType || 'survivor'; // Default to survivor
  const sheetName = sheetType.toUpperCase();

  let sheet = ss.getSheetByName(sheetName);
  let fresh = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    fresh = true;
  }

  sheet.setTabColor(survElimTabColors[sheetType]);

  const totalMembers = memberData.memberOrder.length;
  const members = memberData.memberOrder.map(id => [memberData.members[id]?.name]);

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();

  let previousDataRange, previousData;
  if (!fresh){
    previousDataRange = sheet.getRange(2,3,maxRows-2,WEEKS - WEEKS_TO_EXCLUDE.length);
    previousData = previousDataRange.getValues();
    const text = `💾 Gathered previous data for ${sheetName} sheet, recreating sheet now`;
    Logger.log(text);
    ss.toast(text,`${sheetName} BACKED UP`);
  }
  sheet.clear();

  let rows = totalMembers + 2;
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let cols = WEEKS - WEEKS_TO_EXCLUDE.length + 2;
  if (cols < maxCols) {
    sheet.deleteColumns(cols + 1,maxCols-cols);
  } else if (cols > maxCols) {
    sheet.insertColumnsAfter(maxCols,cols-maxCols);
  }
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,1).setValue('PLAYER');
  let statusCol = 2;
  sheet.getRange(1,statusCol).setValue('STATUS');
  sheet.setColumnWidth(statusCol,80);
  let livesCol = 3;
  sheet.getRange(1,livesCol).setValue('LIVES');
  sheet.setColumnWidth(livesCol,80);
  let revivesCol = 4;
  sheet.getRange(1,revivesCol).setValue('REVIVES');
  sheet.setColumnWidth(revivesCol,65);
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  
  for (let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+5).setValue(weeks[a]);
    sheet.setColumnWidth(a+5,30);
  }

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members);
  sheet.getRange(totalMembers+2,1).setValue('REMAINING');
  sheet.getRange(1,1,totalMembers+2,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  range = sheet.getRange(totalMembers+2,1,1,maxCols);
  range.setBackground('#e6e6e6');
  
  sheet.setFrozenColumns(4);
  sheet.setFrozenRows(1);
  
  ss.setNamedRange(`${sheetName}_NAMES`,sheet.getRange(2,1,totalMembers,1));
  ss.setNamedRange(`${sheetName}_LIVES`,sheet.getRange(2,2,totalMembers,1))
  ss.setNamedRange(`${sheetName}_REVIVES`,sheet.getRange(2,3,totalMembers,1))
  ss.setNamedRange(`${sheetName}_ELIMINATED`,sheet.getRange(2,4,totalMembers,1))
  ss.setNamedRange(`${sheetName}_PICKS`,sheet.getRange(2,5,totalMembers,weeks.length));

  // if (config[`${sheetType}Lives`] == 1) sheet.hideColumns(livesCol);
  if (!config[`${sheetType}Revives`]) sheet.hideColumns(revivesCol);

  if (!fresh) {
    previousDataRange.setValues(previousData);
    const text = `🔄 Previous values restored for ${sheetName} sheet if they were present`;
    Logger.log(text);
    ss.toast(text,`${sheetName} RESTORED`);
  }
  return sheet;
}

// WINNERS Sheet Creation / Adjustment
function winnersSheet(ss,year) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'WINNERS';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  
  let checkboxRange = sheet.getRange(2,3,WEEKS+3,1);
  let checkboxes = checkboxRange.getValues();
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  let rows = weeks.length + 5;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( 3 < maxCols ) {
    sheet.deleteColumns(3,maxCols-3);
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue(year);
  sheet.getRange(1,2).setValue('WINNER');
  sheet.getRange(1,3).setValue('PAID');

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,2,rows-1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,80);
  sheet.setColumnWidth(2,150);
  sheet.setColumnWidth(3,40);

  range = sheet.getRange(2,3,rows-1,1);
  range.insertCheckboxes();
  range.setHorizontalAlignment('center');
  range = sheet.getRange(1,1,rows,2);
  range.setHorizontalAlignment('left');
  let a = 0;
  for (a; a <= weeks.length; a++) {
    sheet.getRange(a+2,1,1,1).setValue(weeks[a]);
  }
  sheet.getRange(a+1,1,4,1).setValues([['SURVIVOR'],['ELIMINATOR'],['MNF'],['OVERALL']]);

  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenRows(1); 

  range = sheet.getRange('R2C2:R'+(weeks.length+1)+'C2');
  ss.setNamedRange('WEEKLY_WINNERS',range);

  sheet.clearConditionalFormatRules(); 
  // OVERALL SHEET GRADIENT RULE
  let fivePlusWins = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=countif($2:B$'+(weeks.length+1)+',B2)>=5')
  .setBackground('#2CFF75')
  .setRanges([range])
  .build();
  let fourPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=4')
    .setBackground('#72FFA3')
    .setRanges([range])
    .build();
  let threePlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=3')
    .setBackground('#BBFFD3')
    .setRanges([range])
    .build();
  let twoPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=2')
    .setBackground('#D3FFE2')
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(fivePlusWins);
  formatRules.push(fourPlusWins);
  formatRules.push(threePlusWins);
  formatRules.push(twoPlusWins);
  sheet.setConditionalFormatRules(formatRules);
  
  // Rewrites the checkboxes if they previously had any checked.
  let col = checkboxRange.getColumn();
  for (let a = 0; (a < checkboxes.length || a < (weeks.length + 4)); a++) {
    if (checkboxes[a][0]) {
      sheet.getRange(a+1,col).check();
    }
  }
  let winRange;
  let nameRange;

  for ( let b = 1; b <= weeks.length; b++ ) {
    winRange = 'WIN_' + (b);
    nameRange = 'NAMES_' + (b);
    sheet.getRange(b+1,2,1,1).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }

  return sheet;

}

// SUMMARY Sheet Creation / Adjustment
function summarySheet(ss,memberData,config) {
  ss = fetchSpreadsheet(ss);
  let restoreNotes = false;
  let notesRange, notes, sheetName = 'SUMMARY';
  
  let docProps;
  if (!memberData || !config) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;
  
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};

  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  } else {
    restoreNotes = true;
    notesRange = sheet.getRange(2,sheet.getRange(1,1,sheet.getMaxRows()-1,sheet.getMaxColumns()).getValues().flat().indexOf('NOTES')+1,sheet.getMaxRows()-1,1);
    notes = notesRange.getValues();
  }
  sheet.clear();
  sheet.setTabColor(winnersTabColor);

  let headers = ['PLAYER'];
  let namedRanges = ['NAME'];
  let headersWidth = [120];
  let mnfCol;
  if (config.pickemsInclude) {
    headers = headers.concat(['TOTAL CORRECT','TOTAL RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS']);
    namedRanges = namedRanges.concat(['PICKS','RANK','AVG_PCT','AVG_PCT_RANK','WINS']);
    headersWidth = headersWidth.concat([90,90,90,90,90]);
    if (!config.mnfExclude) {
      headers = headers.concat(['MNF CORRECT','MNF RANK']);
      namedRanges = namedRanges.concat(['MNF','MNF_RANK']);
      headersWidth = headersWidth.concat([90,90]);
      mnfCol = headers.indexOf('MNF CORRECT') + 1;
    }
  }

  let survivorCol,eliminatorCol
  if (config.survivorInclude) {
    headers.push('SURVIVOR LIVES');
    headers.push('SURVIVOR STATUS');
    namedRanges.push('SURVIVOR_LIVES');
    namedRanges.push('SURVIVOR_STATUS');
    headersWidth.push(90);
    headersWidth.push(90);
    survivorCol = headers.indexOf('SURVIVOR STATUS')+1;
  }
  if (config.eliminatorInclude) {
    headers.push('ELIMINATOR LIVES');
    headers.push('ELIMINATOR STATUS');
    namedRanges.push('ELIMINATOR_LIVES');
    namedRanges.push('ELIMINATOR_STATUS');
    headersWidth.push(100);
    headersWidth.push(100);
    eliminatorCol = headers.indexOf('ELIMINATOR STATUS')+1;
  }
  headers.push('NOTES');
  headersWidth.push(160);
  
  let totalCol = headers.indexOf('TOTAL CORRECT') + 1;
  let weeklyPercentCol = headers.indexOf('AVG % CORRECT') + 1;
  let weeklyRankAvgCol = headers.indexOf('AVG % CORRECT RANK') + 1;
  let weeklyWinsCol = headers.indexOf('WEEKLY WINS') + 1;
  let notesCol = headers.indexOf('NOTES') + 1;

  let len = headers.length;
  
  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( len < maxCols ) {
    sheet.deleteColumns(len,maxCols-len);
  } else if ( len > maxCols ) {
    sheet.insertColumnsAfter(maxCols, len - maxCols);
  }
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,1,1,len).setValues([headers]);
  if(restoreNotes) {
    sheet.getRange(2,notesCol,notes.length,1).setValues(notes);
  }
  
  for ( let a = 0; a < len; a++ ) {
    sheet.setColumnWidth(a+1,headersWidth[a]);
  }
  sheet.setRowHeight(1,40);
  let range = sheet.getRange(1,1,1,maxCols);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  range = sheet.getRange(1,1,rows,len);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  
  range = sheet.getRange(1,1,1,len);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);
  
  sheet.clearConditionalFormatRules(); 
  let formatRules = sheet.getConditionalFormatRules();
  if (config.pickemsInclude) {
    // SUMMARY TOTAL GRADIENT RULE
    let rangeSummaryTot = sheet.getRange('R2C'+totalCol+':R'+rows+'C'+totalCol);
    let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#75F0A1')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([rangeSummaryTot])
      .build();
    formatRules.push(formatRuleOverallTot);
    // MNF TOTAL GRADIENT RULES
    let rangeMNFTot, rangeMNFRank, formatRuleMNFRank;
    if (!config.mnfExclude) {
      rangeMNFTot = sheet.getRange('R2C'+mnfCol+':R'+rows+'C'+mnfCol);
      let formatRuleMNFTot = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint('#75F0A1')
        .setGradientMinpoint('#FFFFFF')
        .setRanges([rangeMNFTot])
        .build();
      formatRules.push(formatRuleMNFTot);    
      // RANK MNF GRADIENT RULE
      rangeMNFRank = sheet.getRange('R2C'+(mnfCol+1)+':R'+rows+'C'+(mnfCol+1));
      formatRuleMNFRank = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
        .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
        .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
        .setRanges([rangeMNFRank])
        .build();
      formatRules.push(formatRuleMNFRank);
    }
    // RANK OVERALL RULE
    let rangeOverallRank = sheet.getRange('R2C'+(totalCol+1)+':R'+rows+'C'+(totalCol+1));
    let formatRuleRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
      .setRanges([rangeOverallRank])
      .build();
    formatRules.push(formatRuleRank);
    // WEEKLY WINS GRADIENT/SINGLE COLOR RULES
    range = sheet.getRange('R2C'+weeklyWinsCol+':R'+rows+'C'+weeklyWinsCol);
    let formatRuleWeeklyWinsEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#FFFFFF')
      .setFontColor('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWinsEmpty);
    let formatRuleWeeklyWins = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#ffee00')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWins);   
    // OVERALL AND WEEKLY CORRECT % AVG
    range = sheet.getRange('R2C'+weeklyPercentCol+':R'+rows+'C'+weeklyPercentCol);
    range.setNumberFormat('##.#%');
    let formatRuleCorrectAvg = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
      .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectAvg);
    // WEEKLY RANK AVG
    range = sheet.getRange('R2C'+weeklyRankAvgCol+':R'+rows+'C'+weeklyRankAvgCol);
    range.setNumberFormat('#.#');
    let formatRuleCorrectRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, "5")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "15")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectRank);
  }
  if (config.survivorInclude) {
  // SURVIVOR "IN"
    range = sheet.getRange('R2C'+survivorCol+':R'+(totalMembers+1)+'C'+survivorCol);
    let formatRuleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#C9FFDF')
      .setRanges([range])
      .build();
    // SURVIVOR "OUT"
    let formatRuleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OUT')
      .setBackground('#F2BDC2')
      .setRanges([range])
      .build();    
    formatRules.push(formatRuleIn);
    formatRules.push(formatRuleOut);
  }
  if (config.eliminatorInclude) {
  // ELIMINATOR "IN"
    range = sheet.getRange('R2C'+eliminatorCol+':R'+(totalMembers+1)+'C'+eliminatorCol);
    let formatRuleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#C9FFDF')
      .setRanges([range])
      .build();
    // ELIMINATOR "OUT"
    let formatRuleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OUT')
      .setBackground('#F2BDC2')
      .setRanges([range])
      .build();    
    formatRules.push(formatRuleIn);
    formatRules.push(formatRuleOut);
  }  
  sheet.setConditionalFormatRules(formatRules);
  // Loops through and defines all named ranges for Summary Sheet (TOT is prefixed)
  namedRanges.forEach((name,index) => {
    ss.setNamedRange(`TOT_${name}`,sheet.getRange(2,index+1,totalMembers,1));
  });
  
  // Creates all formulas for SUMMARY Sheet
  summarySheetFormulas(headers, sheet,totalMembers,ss);

  return sheet;  
}

// LEADERBOARD Sheet Creation
function leaderboardSheet(ss, config, memberData) {
  ss = ss || fetchSpreadsheet(ss);
  const summaryAvailable = ss.getSheetByName('SUMMARY') != null;
  if (!summaryAvailable) Logger.log(`⚠️ No SUMMARY sheet found: no season-long metrics will be displayed on the sheet until this is created. Likely needs the leaderboardSheet to be re-run...`)
  let docProps = (!config || !memberData) ? PropertiesService.getDocumentProperties() : null;
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};

  const totalMembers = memberData.memberOrder ? memberData.memberOrder.length : 0;
  if (totalMembers <= 0) {
    SpreadsheetApp.getUi().alert('⚠️ MEMBER ISSUE', 'Please populate members before building the leaderboard.', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }

  const sheetName = 'LEADERBOARD';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
    sheet.clearNotes();
    sheet.clearConditionalFormatRules();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  } else {
    sheet = ss.insertSheet(sheetName, 1);
  }
  sheet.setTabColor(leaderboardTabColor);

  const isAts = !!config.pickemsAts;
  let formatRules = [];
  
  // -------------------------------------------------------------
  // 1. ROW DEFINITIONS & LAYOUT
  // -------------------------------------------------------------
  const firstRow         = 1; // Controls: Week, Sort, Contenders Filter
  const subHeaderRow     = 2; // Game Status (✅/⏳) & Weekday
  const weekDayRow       = 3; // Weekdays for the matchup block, merged for others
  const matchupRow       = 4; // Main Away @ Home Matchups & Column Emojis
  const outcomeRow       = 5; // Straight-Up Winner Outcome
  const spreadOutcomeRow = 6; // ATS Winner Outcome
  const summaryRow       = 7; // Static Group Stats Row
  const dataStartRow     = 8; // First Player Row
  const dataEndRow       = dataStartRow + totalMembers - 1;
  const freezeRow        = summaryRow;

  // -------------------------------------------------------------
  // 2. COLUMN DEFINITIONS
  // -------------------------------------------------------------
  let colHeaders = ['PLAYER'];
  let colSubHeaders = ['PLAYER'];
  let colWidths  = [140];
  let fontSizes  = [11];
  let subHeaderFontSizes = [7];
  let survCol, elimCol;
  
  // A. Season & Pool Totals (From SUMMARY Sheet)
  const overallStartCol = 2;
  if (config.pickemsInclude) {
    colHeaders.push('⭐', '🥇');
    colSubHeaders.push('Points Tot','Rank');
    colWidths.push(60, 60);
    fontSizes.push(11, 11);
    subHeaderFontSizes.push(7,7);
    if (!config.mnfExclude) {
      colWidths.push(60);
      colHeaders.push('🌙');
      colSubHeaders.push('MNF Tot')
      fontSizes.push(11);
    }
  }
  if (config.survivorInclude) {
    survCol = colHeaders.length + 1;
    colWidths.push(65, 75);
    colHeaders.push('👑', '👑');
    colSubHeaders.push('Lives','Status');
    fontSizes.push(11, 11);
    subHeaderFontSizes.push(7,7);
  }
  if (config.eliminatorInclude) {
    elimCol = colHeaders.length + 1;
    colWidths.push(65, 75);
    colHeaders.push('💀', '💀');
    colSubHeaders.push('Lives','Status');
    fontSizes.push(11, 11);
    subHeaderFontSizes.push(7,7);
  }
  const overallEndCol = colHeaders.length;

  // B. Active Weekly Performance
  const weeklyStartCol = overallEndCol + 1;
  colHeaders.push('⭐', '🥇', '💯', '🎲', '📊', '🃏'); // Points, Rank, %, Chances, Sparkline, Wildcard
  colSubHeaders.push('Picks','Rank','Percent','Chances','Chances','Wildcard');
  colWidths.push(50, 50, 50, 55, 60, 50);
  fontSizes.push(14, 14, 14, 14, 14, 14);
  subHeaderFontSizes.push(7, 7, 7, 7, 7, 7);

  let tiebreakerCol = -1;
  if (config.tiebreakerInclude) {
    colWidths.push(50, 50);
    colHeaders.push('⚖️', '📏');
    colSubHeaders.push('Tiebreaker','Difference');
    fontSizes.push(14, 14);
    subHeaderFontSizes.push(7,7);
    tiebreakerCol = colHeaders.length - 1;
  }

  let commentCol = -1;
  if (!config.commentsExclude) {
    colWidths.push(150);
    colHeaders.push('💬');
    colSubHeaders.push('Comments');
    fontSizes.push(14);
    subHeaderFontSizes.push(7);
    commentCol = colHeaders.length;
  }
  const weeklyEndCol = colHeaders.length;

  // C. Matchup Columns (Max 16 games)
  const firstMatchupCol = weeklyEndCol + 1;
  const maxWeeklyGames = 16;
  for (let g = 1; g <= maxWeeklyGames; g++) {
    colHeaders.push(`G${g}`);
    colWidths.push(52);
    fontSizes.push(9);
  }
  const finalMatchupCol = colHeaders.length;

  adjustRows(sheet, dataEndRow + 2);
  adjustColumns(sheet, finalMatchupCol);

  for (let c = 0; c < colWidths.length; c++) {
    sheet.setColumnWidth(c + 1, colWidths[c]);
  }
  const fullSheetRange = sheet.getRange(1, 1, dataEndRow + 2, finalMatchupCol);
  fullSheetRange.setFontFamily("Montserrat").setVerticalAlignment("middle");

  // -------------------------------------------------------------
  // 3. CONTROL PANEL (Row 1)
  // -------------------------------------------------------------
  sheet.setRowHeight(firstRow, 34);
  sheet.getRange(firstRow, 1, 1, finalMatchupCol).setBackground('#1E1E1E').setFontColor('#FFFFFF');

  // Week Selector (B1)
  sheet.getRange(firstRow, 1).setValue('WEEK').setFontWeight('bold').setHorizontalAlignment('right').setFontSize(16);
  const weekCell = sheet.getRange(firstRow, 2);
  const weekCellString = weekCell.getA1Notation();
  const weekCellStringRC = `R${firstRow}C2`;
  weekCell.setValue(1).setFontWeight('bold').setHorizontalAlignment('center').setFontSize(16);
  const weekRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Array.from({ length: 18 }, (_, i) => `${i + 1}`), true)
    .build();
  weekCell.setDataValidation(weekRule);

  // Sort Selector (D1)
  sheet.getRange(firstRow, 3).setValue('Sort').setFontWeight('bold').setHorizontalAlignment('right').setFontSize(10);
  const sortCell = sheet.getRange(firstRow, 4);
  sortCell.setValue('Rank').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(10);
  const sortRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Rank', 'Chances', 'Default'], true)
    .build();
  sortCell.setDataValidation(sortRule);

  // Hide 0% Chances Checkbox (F1)
  sheet.getRange(firstRow, 5).setValue('Only Contenders').setFontWeight('bold').setHorizontalAlignment('right').setFontSize(10);
  const filterCell = sheet.getRange(firstRow, 6);
  filterCell.insertCheckboxes().setHorizontalAlignment('center').setBackground('#000000').setFontColor('#ffffff').setFontSize(14);

  // -------------------------------------------------------------
  // 4. SECTION HEADERS & GAME PROGRESS BAR (Row 2)
  // -------------------------------------------------------------
  sheet.getRange(subHeaderRow,1).setValue('MEMBERS').setBackground('#333333').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  if (overallEndCol >= overallStartCol) {
    sheet.getRange(subHeaderRow, overallStartCol, 1, overallEndCol - overallStartCol + 1).merge()
         .setValue('SEASON & POOL TOTALS').setBackground('#263238').setFontColor('#ECEFF1').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);
  }

  // Active Weekly Performance Header with Embedded Game Progress
  const totalGamesFormula = `COUNTA(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol})`;
  const completedGamesFormula = `COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})`;
  
  sheet.getRange(firstRow, weeklyStartCol, 1, weeklyEndCol - weeklyStartCol).merge()
      .setValue(`WEEKLY METRICS`).setBackground('#1B5E20').setFontColor('#E8F5E9').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);
  
  sheet.getRange(subHeaderRow, weeklyStartCol, 2, 2).merge()
       .setValue('="GAME"&char(10)&"PROGRESS"').setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('right').setFontSize(9);
  
  const progressTextFormula = `=IFERROR(${completedGamesFormula} & " / " & ${totalGamesFormula} & " (" & TEXT(${completedGamesFormula}/${totalGamesFormula},"0%") & ")", "0%")`;
  sheet.getRange(subHeaderRow, weeklyStartCol + 2, 2, 2).merge()
       .setFormulaR1C1(progressTextFormula).setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);

  const progressSparklineFormula = `=IFERROR(SPARKLINE(${completedGamesFormula}, {"charttype","bar";"max",${totalGamesFormula};"color1","#00E676";"color2","#424242"}),"")`;
  sheet.getRange(subHeaderRow, weeklyStartCol + 4, 2, weeklyEndCol - (weeklyStartCol + 4) + 1).merge()
       .setFormulaR1C1(progressSparklineFormula).setBackground('#1B5E20');

  // Matchups Super-Header
  sheet.getRange(firstRow, firstMatchupCol, 1, maxWeeklyGames).merge()
       .setValue('WEEKLY MATCHUP PICKS').setBackground('#0D47A1').setFontColor('#E3F2FD').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(12);

  // -------------------------------------------------------------
  // 5. MATCHUPS, STATUS & OUTCOMES VIA NAMED RANGES (Rows 3-6)
  // -------------------------------------------------------------
  sheet.setRowHeight(matchupRow, 38);
  sheet.setRowHeight(outcomeRow, 24);
  sheet.setRowHeight(spreadOutcomeRow, 24);
  sheet.setRowHeight(summaryRow, 28);

  // Static Column Emojis / Names
  sheet.getRange(matchupRow, 1, 1, weeklyEndCol).setValues([colHeaders.slice(0, weeklyEndCol)])
       .setBackground('#000000').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center');
  sheet.getRange(outcomeRow, 1, 1, weeklyEndCol).setValues([colSubHeaders.slice(0, weeklyEndCol)])
       .setBackground('#000000').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(7).setHorizontalAlignment('center');

  for (let c = 0; c < weeklyEndCol; c++) {
    sheet.getRange(matchupRow, c + 1).setFontSize(fontSizes[c]);
  }

  // Row 4: Matchup Names (Spills from LEAGUE_week)
  sheet.getRange(matchupRow, firstMatchupCol).setFormula(
    `=IFERROR(INDIRECT("${LEAGUE}_" & ${weekCellString}),"")`
  );
  sheet.getRange(matchupRow, firstMatchupCol, 1, maxWeeklyGames)
       .setBackground('#000000').setFontColor('#FFFFFF').setFontSize(9).setFontWeight('bold')
       .setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Row 5: Winner (straight up) for Picks (Spills from LEAGUE_PICKEM_OUTCOMES_week)
  sheet.getRange(outcomeRow, 1).setValue('WINNER').setFontWeight('bold').setFontSize(9).setBackground('#CFD8DC').setHorizontalAlignment('left');
  sheet.getRange(outcomeRow, firstMatchupCol).setFormula(
    `=IFERROR(INDIRECT("${LEAGUE}_PICKEM_OUTCOMES_" & ${weekCellString}),"")`
  );
  sheet.getRange(outcomeRow, firstMatchupCol, 1, maxWeeklyGames)
       .setBackground('#ECEFF1').setFontWeight('bold').setHorizontalAlignment('center');

  // Row 6: Winner ATS (Spills from LEAGUE_ATS_OUTCOMES_week)
  sheet.getRange(spreadOutcomeRow, 1).setValue('WINNER (ATS)').setFontWeight('bold').setFontSize(9).setBackground('#CFD8DC').setHorizontalAlignment('left');
  sheet.getRange(spreadOutcomeRow, firstMatchupCol).setFormula(
    `=IFERROR(INDIRECT("${LEAGUE}_ATS_OUTCOMES_" & ${weekCellString}),"")`
  );
  sheet.getRange(spreadOutcomeRow, firstMatchupCol, 1, maxWeeklyGames)
       .setBackground('#ECEFF1').setFontWeight('bold').setHorizontalAlignment('center');

  if (isAts) {
    sheet.hideRows(outcomeRow);
  } else {
    sheet.hideRows(spreadOutcomeRow);
  }

  // Row 2: Status (✅ / ⏳)
  const effectiveOutcomeRow = isAts ? spreadOutcomeRow : outcomeRow;
  for (let g = 0; g < maxWeeklyGames; g++) {
    const colIdx = firstMatchupCol + g;
    sheet.getRange(subHeaderRow, colIdx).setFormulaR1C1(
      `=iferror(IF(ISBLANK(R${matchupRow}C[0]),"",IF(NOT(ISBLANK(R${effectiveOutcomeRow}C[0])),"✅","⏳")),"")`
    ).setHorizontalAlignment('center').setBackground('#0D47A1').setFontColor('#E3F2FD').setFontSize(8).setFontWeight('bold');
  }
  // Row 3: Weekday Day Name
  sheet.getRange(weekDayRow,firstMatchupCol).setFormula(`=iferror(indirect("${LEAGUE}_DAYS_"&${weekCellString}),"")`);
  sheet.getRange(weekDayRow,firstMatchupCol,1,MAXGAMES).setBackground('#000000').setFontColor('#FFFFFF').setHorizontalAlignment('center').setFontSize(7);
  
  // Row 4: Matchups 
  sheet.getRange(matchupRow,firstMatchupCol).setFormula(`=iferror(indirect("${LEAGUE}_"&${weekCellString}),"")`);
  sheet.getRange(matchupRow,firstMatchupCol,1,MAXGAMES).setHorizontalAlignment('center').setFontSize(8).setFontWeight('bold');
  
  // Row 7: Home/Away Splits Across Matchups
  sheet.getRange(summaryRow,firstMatchupCol).setFormula(`=iferror(indirect("${LEAGUE}_BIAS_" & ${weekCellString}),"")`);
  
  // -------------------------------------------------------------
  // 6. STATIC GROUP STATS SUMMARY ROW (Row 7)
  // -------------------------------------------------------------
  sheet.getRange(summaryRow, 1, 1, finalMatchupCol).setBackground('#ECEFF1').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(summaryRow, 1).setValue('Group Stats').setHorizontalAlignment('left').setFontSize(9);

  const pointsColIdx = weeklyStartCol;
  const rankColIdx   = weeklyStartCol + 1;
  const pctColIdx    = weeklyStartCol + 2;
  const wildColIdx   = weeklyStartCol + 5;

  sheet.getRange(summaryRow, pointsColIdx).setFormulaR1C1(
    `=IFERROR(IF(SUM(R${dataStartRow}C[0]:R${dataEndRow}C[0])>0, "AVG:"&CHAR(10)&TEXT(ROUND(AVERAGE(R${dataStartRow}C[0]:R${dataEndRow}C[0]),1),"#.0"),""),"")`
  );
  sheet.getRange(summaryRow, rankColIdx).setFormulaR1C1(
    `=IFERROR(IF(COUNTIF(R${dataStartRow}C[0]:R${dataEndRow}C[0],1)>1, COUNTIF(R${dataStartRow}C[0]:R${dataEndRow}C[0],1)&"-Way"&CHAR(10)&"Tie", "Leader:"&CHAR(10)&INDEX(R${dataStartRow}C1:R${dataEndRow}C1, MATCH(1, R${dataStartRow}C[0]:R${dataEndRow}C[0], 0))),"")`
  );
  sheet.getRange(summaryRow, pctColIdx).setFormulaR1C1(
    `=IFERROR(IF(COUNTA(R${dataStartRow}C[0]:R${dataEndRow}C[0])>0, TEXT(AVERAGE(R${dataStartRow}C[0]:R${dataEndRow}C[0]), "0.0%"),""),"")`
  );
  sheet.getRange(summaryRow, wildColIdx).setFormulaR1C1(
    `=IFERROR(IF(SUM(R${dataStartRow}C[0]:R${dataEndRow}C[0])>0, TEXT(AVERAGE(R${dataStartRow}C[0]:R${dataEndRow}C[0]), "0.0%"),""),"")`
  );

  if (config.tiebreakerInclude) {
    sheet.getRange(summaryRow, tiebreakerCol).setFormulaR1C1(`=IFERROR("AVG:"&CHAR(10)&ROUND(AVERAGE(R${dataStartRow}C[0]:R${dataEndRow}C[0]),1),"")`);
    sheet.getRange(summaryRow, tiebreakerCol + 1).setFormulaR1C1(`=IFERROR("AVG:"&CHAR(10)&ROUND(AVERAGE(R${dataStartRow}C[0]:R${dataEndRow}C[0]),1),"")`);
  }

  // BIAS COLOR SCHEMES (copy from weeklySheet())
  let homeAwayPercents = [90,80,70,60,50];
  let awayColors = ['#FFFB7D','#FFFC96','#FFFCB0','#FFFDC9','#FFFEE3'];
  let homeColors = ['#7DFFFB','#96FFFC','#B0FFFC','#C9FFFD','#E3FFFE'];
  let awayFormula = `=and(regexextract(indirect("R[0]C[0]",false),"[A-Z]{2,3}")=regexextract(indirect("R${matchupRow}C[0]",false),"[A-Z]{2,3}"),value(regexextract(indirect("R[0]C[0]",false),"[0-9\.]+"))>=%%)`; // Replaceable "%%" for inserting percent number
  let homeFormula = `=and(regexextract(indirect("R[0]C[0]",false),"[A-Z]{2,3}")=regexextract(right(indirect("R${matchupRow}C[0]",false),3),"[A-Z]{2,3}"),value(regexextract(indirect(\"R[0]C[0]",false),"[0-9\.]+"))>=%%)`; // Replaceable "%%" for inserting percent number
  const preferenceRange = sheet.getRange(summaryRow,firstMatchupCol,1,MAXGAMES); // Summary row of matchups
  for (let a = 0; a < homeAwayPercents.length; a++) {
    let formula = awayFormula.replace('%%',homeAwayPercents[a]);

    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(awayColors[a])
      .setRanges([preferenceRange]);
    rule.build();
    formatRules.push(rule);

    formula = homeFormula.replace('%%',homeAwayPercents[a]);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(homeColors[a])
      .setRanges([preferenceRange]);
    rule.build();
    formatRules.push(rule);    
  }

  // -------------------------------------------------------------
  // 7. COLUMN A: SORTED & FILTERED PLAYER NAMES
  // -------------------------------------------------------------
  // Pulls NAMES_week, RNK_week, CHANCES_week and applies sort & contenders filter
  const colAFormula = `=LET(
    names, INDIRECT("NAMES_" & ${weekCellString}),
    ranks, IFERROR(INDIRECT("RNK_" & ${weekCellString}), 999),
    chances, IFERROR(INDIRECT("CHANCES_" & ${weekCellString}), 0),
    hideZero, $F$1,

    filtered, FILTER(
      {names, ranks, chances, SEQUENCE(ROWS(names), 1, 1, 1)},
      IF(hideZero, chances > 0, ROW(names) > 0)
    ),

    fNames, INDEX(filtered,,1),
    fRanks, INDEX(filtered,,2),
    fChances, INDEX(filtered,,3),
    fSeq, INDEX(filtered,,4),

    sortKey, IFS($D$1="Rank", fRanks, $D$1="Chances", -fChances, $D$1="Default", fSeq, TRUE, fSeq),

    IFERROR(SORT(fNames, sortKey, 1), names)
  )`;

  sheet.getRange(dataStartRow, 1).setFormula(colAFormula);

  // -------------------------------------------------------------
  // 8. DATA LOOKUPS VIA XLOOKUP & NAMED RANGES
  // -------------------------------------------------------------
  for (let r = dataStartRow; r <= dataEndRow; r++) {
    const playerCell = `R${r}C1`;
  
    // A. Season & Pool Totals (XLOOKUP into SUMMARY sheet)
    if (config.pickemsInclude) {
      sheet.getRange(r, 2).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},SUMMARY_NAMES,SUMMARY!$B$2:$B)),"")`); // TOT ⭐
      sheet.getRange(r, 3).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},SUMMARY_NAMES,SUMMARY!$C$2:$C)),"")`); // TOT 🥇
      if (!config.mnfExclude) {
        sheet.getRange(r, 4).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},SUMMARY_NAMES,TOT_MNF_CORRECT,))),"")`);
      }
    }
    if (config.survivorInclude) {
      sheet.getRange(r, survCol).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, SUMMARY_NAMES, INDEX(SUMMARY!$A$2:$Z, 0, MATCH("SURVIVOR LIVES", SUMMARY!$1:$1, 0)))),"")`);
      sheet.getRange(r, survCol + 1).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, SUMMARY_NAMES, INDEX(SUMMARY!$A$2:$Z, 0, MATCH("SURVIVOR STATUS", SUMMARY!$1:$1, 0)))),"")`);
    }
    if (config.eliminatorInclude) {
      sheet.getRange(r, elimCol).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, SUMMARY_NAMES, INDEX(SUMMARY!$A$2:$Z, 0, MATCH("ELIMINATOR LIVES", SUMMARY!$1:$1, 0)))),"")`);
      sheet.getRange(r, elimCol + 1).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, SUMMARY_NAMES, INDEX(SUMMARY!$A$2:$Z, 0, MATCH("ELIMINATOR STATUS", SUMMARY!$1:$1, 0)))),"")`);
    }

    // B. Weekly Performance Lookups (XLOOKUP into active week's named ranges)
    sheet.getRange(r, weeklyStartCol).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},INDIRECT("NAMES_" & ${weekCellStringRC}),INDIRECT("TOT_"&${weekCellStringRC}))),"")`);
    sheet.getRange(r, weeklyStartCol + 1).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},INDIRECT("NAMES_" & ${weekCellStringRC}),INDIRECT("RNK_"&${weekCellStringRC}))),"")`);
    sheet.getRange(r, weeklyStartCol + 2).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},INDIRECT("NAMES_" & ${weekCellStringRC}),INDIRECT("PCT_"&${weekCellStringRC}))),"")`);
    sheet.getRange(r, weeklyStartCol + 3).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","",XLOOKUP(${playerCell},INDIRECT("NAMES_" & ${weekCellStringRC}),INDIRECT("CHANCES_"&${weekCellStringRC}))),"")`);

    // Sparkline
    sheet.getRange(r, weeklyStartCol + 4).setFormulaR1C1(
      `=IFERROR(IF(OR(${playerCell}="", ISBLANK(R[0]C[-1])),"", SPARKLINE(MAX(R[0]C[-1], 0.05), {"charttype","bar";"max",1;"color1", IF(R[0]C[-1]=MAX(R${dataStartRow}C[-1]:R${dataEndRow}C[-1]), "#00E676", IF(R[0]C[-1]<(MAX(R${dataStartRow}C[-1]:R${dataEndRow}C[-1])/3), "#FF8A80", "#FFD54F"))})),"")`
    );

    // Wildcard
    sheet.getRange(r, weeklyStartCol + 5).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, INDIRECT("NAMES_" & ${weekCellStringRC}), INDIRECT("WILDCARD_" & ${weekCellStringRC}))),"")`);

    // Tiebreaker & Diff
    if (config.tiebreakerInclude) {
      sheet.getRange(r, tiebreakerCol).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, INDIRECT("NAMES_" & ${weekCellStringRC}), INDIRECT("${LEAGUE}_TIEBREAKER_" & ${weekCellStringRC}))),"")`);
      sheet.getRange(r, tiebreakerCol + 1).setFormulaR1C1(`=IFERROR(IF(OR(${playerCell}="",ISBLANK(R[0]C[-1]),ISBLANK(INDIRECT("${LEAGUE}_TIEBREAKER_"&${weekCellStringRC}&"_OUTCOME"))),"",ABS(R[0]C[-1]-INDIRECT("${LEAGUE}_TIEBREAKER_"&${weekCellStringRC}&"_OUTCOME"))),"")`);
    }

    // Comments
    if (!config.commentsExclude) {
      sheet.getRange(r, commentCol).setFormulaR1C1(`=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, INDIRECT("NAMES_" & ${weekCellStringRC}), INDIRECT("COMMENTS_" & ${weekCellStringRC}))),"")`);
    }

    // C. Matchup Picks Array Spill: XLOOKUP into 2D picks matrix (Spills entire row of picks)
    sheet.getRange(r, firstMatchupCol).setFormulaR1C1(
      `=IFERROR(IF(${playerCell}="","", XLOOKUP(${playerCell}, INDIRECT("NAMES_" & ${weekCellStringRC}), INDIRECT("${LEAGUE}_PICKS_" & ${weekCellStringRC}))),"")`
    );
  }

  // Formats
  sheet.getRange(dataStartRow, weeklyStartCol + 2, totalMembers, 1).setNumberFormat('0.0%');
  sheet.getRange(dataStartRow, weeklyStartCol + 3, totalMembers, 1).setNumberFormat('0.0%');
  sheet.getRange(dataStartRow, weeklyStartCol + 5, totalMembers, 1).setNumberFormat('0.0%');

  // -------------------------------------------------------------
  // 9. BORDERS & FREEZE
  // -------------------------------------------------------------
  // sheet.setFrozenRows(freezeRow);
  sheet.setFrozenColumns(1);

  const sectionColsToBorder = [1, overallEndCol, weeklyEndCol, finalMatchupCol];
  sectionColsToBorder.forEach(c => {
    sheet.getRange(1, c, dataEndRow, 1).setBorder(null, null, null, true, null, null, '#424242', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.getRange(summaryRow, 1, 1, finalMatchupCol).setBorder(null, null, true, null, null, null, '#424242', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // -------------------------------------------------------------
  // 10. CONDITIONAL FORMATTING (pickColors Gradients)
  // -------------------------------------------------------------
  
  // A. Weekday Backgrounds
  if (typeof dayColorsObj !== 'undefined') {
    const dayRange = sheet.getRange(weekDayRow, firstMatchupCol, 1, MAXGAMES);
    Object.keys(dayColorsObj).forEach(day => {
      formatRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains(day)
          .setBackground(dayColorsObj[day])
          .setFontColor('#000000')
          .setRanges([dayRange])
          .build()
      );
    });
  }

  // B. Completed Matchup Headers Dimming
  const matchupHeadersRange = sheet.getRange(matchupRow, firstMatchupCol, 1, maxWeeklyGames);
  const completedMatchupRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=NOT(ISBLANK(R${effectiveOutcomeRow}C[0]))`)
    .setBackground('#37474F')
    .setFontColor('#80CBC4')
    .setRanges([matchupHeadersRange])
    .build();
  formatRules.push(completedMatchupRule);

  // C. Matchup Grid Picks (Zebra Parity + Bonus Multipliers)
  const bonusCount = 3;
  const parities = {
    even: { fn: 'iseven' },
    odd:  { fn: 'isodd'  }
  };

  const picksGridRange = sheet.getRange(dataStartRow, firstMatchupCol, totalMembers, maxWeeklyGames);
  const outcomeRef = `INDIRECT("R${effectiveOutcomeRow}C[0]", FALSE)`;
  const cellRef    = `INDIRECT("R[0]C[0]", FALSE)`;
  const matchRef   = `INDIRECT("R${matchupRow}C[0]", FALSE)`;
  const bonusRef   = `INDEX(INDIRECT("${LEAGUE}_BONUS_" & ${weekCellString}), 1, COLUMN() - ${firstMatchupCol - 1})`;

  const baseFormulas = {
    correct:   `AND(NOT(ISBLANK(${outcomeRef})), ${cellRef}<>"", ${cellRef}=${outcomeRef})`,
    incorrect: `AND(NOT(ISBLANK(${outcomeRef})), ${cellRef}<>"", ${cellRef}<>${outcomeRef})`,
    home:      `AND(ISBLANK(${outcomeRef}), ${cellRef}<>"", ${cellRef}=TRIM(RIGHT(${matchRef}, LEN(${matchRef})-FIND("@", ${matchRef}))))`,
    away:      `AND(ISBLANK(${outcomeRef}), ${cellRef}<>"", ${cellRef}=TRIM(LEFT(${matchRef}, FIND("@", ${matchRef})-1)))`
  };

  for (const [type, cfg] of Object.entries(pickColors)) {
    for (const parity of Object.values(parities)) {
      const startColor = parity.fn === 'iseven' ? cfg.even : cfg.odd;
      const gradient = typeof hexGradient === 'function'
        ? hexGradient(startColor, cfg.end, bonusCount)
        : [startColor];

      for (let i = gradient.length - 1; i >= 0; i--) {
        const bonusLevel = i + 1;
        const baseFormula = baseFormulas[type];
        const parityCondition = `${parity.fn}(ROW())`;

        let finalFormula = `=AND(${baseFormula}, ${parityCondition}`;
        if (i > 0) {
          finalFormula += `, ${bonusRef}=${bonusLevel}`;
        }
        finalFormula += `)`;

        const ruleBuilder = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(finalFormula)
          .setBackground(gradient[i])
          .setRanges([picksGridRange]);

        if (cfg.font) {
          ruleBuilder.setFontColor(cfg.font);
        }

        formatRules.push(ruleBuilder.build());
      }
    }
  }

  sheet.setConditionalFormatRules(formatRules);

  SpreadsheetApp.flush();
  return sheet;
}

// Helper to look up column index by header text
function getHeaderColIndex(sheet, text) {
  const rowVals = sheet.getRange(4, 1, 1, sheet.getMaxColumns()).getValues()[0];
  const idx = rowVals.indexOf(text);
  return idx > -1 ? idx + 1 : 2;
}

// UPDATES SUMMARY SHEET FORMULAS
function summarySheetFormulas(headers,sheet,totalMembers,ss) {
  let arr = [...headers] || ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS','SURVIVOR LIVES','SURVIVOR STATUS','ELIMINATOR LIVES','ELIMINATOR STATUS','NOTES'];
  
  ss = fetchSpreadsheet(ss);
  if (!sheet) {
    sheet = fetchSpreadsheet().getSheetByName('SUMMARY');  
  }
  headers.unshift('COL INDEX ADJUST');

  for (let a = 0; a < arr.length; a++) {
    for (let b = 0; b < totalMembers; b++) {
      if (headers[a] == 'TOTAL CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_NAMES,TOT_OVERALL},2,false))');
      } else if (headers[a] == 'TOTAL RANK' || headers[a] == 'AVG % CORRECT RANK' || headers[a] == 'MNF RANK') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
      } else if (headers[a] == 'MNF CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{MNF_NAMES,MNF},2,false))');
      } else if (headers[a] == 'AVG % CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_PCT_NAMES,TOT_OVERALL_PCT},2,false))');
      } else if (headers[a] == 'WEEKLY WINS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(countif(WEEKLY_WINNERS,R[0]C1))');
      } else if (headers[a] == 'SURVIVOR STATUS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{SURVIVOR_NAMES,SURVIVOR_ELIMINATED},2,false),)');
      } else if (headers[a] == 'ELIMINATOR STATUS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{ELIMINATOR_NAMES,ELIMINATOR_ELIMINATED},2,false),)');
      }
    }
  }
  Logger.log(`🧮 Updated formulas and ranges for summary sheet`);
}

// TOT / RANK / PCT / MNF Combination formula for sum/average per player row
function overallPrimaryFormulas(sheet,totalMembers,maxCols,action,avgRow) {
  if (action == 'average') {
    sheet.getRange(2,2,totalMembers,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,average(R[0]C3:R[0]C'+maxCols+')))')
      .setNumberFormat("#0.0");
  } else if (action == 'sum') {
    sheet.getRange(2,2,totalMembers,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,sum(R[0]C3:R[0]C'+maxCols+')))')
      .setNumberFormat("##");
  }
  if (sheet.getSheetName() == 'PCT') {
    sheet.getRange(2,2,totalMembers,1).setNumberFormat("##.#%");
  }
  if (avgRow) {
    if (sheet.getSheetName() == 'PCT'){  
      sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
        .setNumberFormat('##.#%');
    } else {
      sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
        .setNumberFormat("#0.0");
    }
  }
}

// TOT / RNK / PCT / MNF Combination formula for each column (week)
function overallMainFormulas(weeks,sheet,totalMembers,str,avgRow) {
  weeks = weeks || Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  for (let a = 0; a < weeks.length; a++ ) {
    for (let b = 1; b <= totalMembers; b++) {
      if (str == 'TOT') {
        sheet.getRange(b+1,a+3).setFormula('=iferror(if(or(iserror(vlookup($A'+(b+1)+',NAMES_'+weeks[a]+',1,false)),counta(filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+'))=0),,arrayformula(countifs(filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+')='+LEAGUE+'_PICKEM_OUTCOMES_'+weeks[a]+',true,filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+'),\"<>\"))),)');
      } else {
        sheet.getRange(b+1,a+3).setFormula('=iferror(arrayformula(vlookup(R[0]C1,{NAMES_'+weeks[a]+','+str+'_'+weeks[a]+'},2,false)))');
      }
      if (sheet.getSheetName() == 'PCT') {
        sheet.getRange(b+1,a+3).setNumberFormat("##.#%");
      } else {
        sheet.getRange(b+1,a+3).setNumberFormat("#0");
      }
    }
  }
  if (avgRow) {
    if (sheet.getSheetName() == 'MNF') {
      // Instance of MNF sheet, where sheet needs to have data for quantity of MNF games
      Logger.log(`📡 Checking for Monday games, if any`);
      let data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(LEAGUE).getValues();
      let text = '0';
      let result = text.repeat(weeks.length);
      let mondayNightGames = Array.from(result);
      for (let a = 0; a < data.length; a++) {
        if ( data[a][2] == 1 && data[a][3] >= 17) {
          mondayNightGames[(data[a][0]-1)]++;
        }
      }
      for (let a = 0; a < weeks.length; a++){
        let rows = sheet.getMaxRows();
        if (mondayNightGames[a] > 1) {
          sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0])/'+mondayNightGames[a]+',))')
            .setNumberFormat("##%");
        } else {
          sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
            .setNumberFormat("##%");
        }
      }
    } else {
      for (let a = 0; a < weeks.length; a++){
        let rows = sheet.getMaxRows();
        sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
          .setNumberFormat("##%");
      }
    }
  }
}

// WEEKLY WINNERS Combination formula update
function winnersFormulas(weeks,sheet) {
  for (let a = 0; a < weeks.length; a++ ) {
    let winRange = `WIN_${weeks[a]}`;
    let nameRange = `NAMES_${weeks[a]}`;
    sheet.getRange(a+1,2).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
}

// REFRESH FORMULAS FOR TOT / RNK / PCT / MNF
function allFormulasUpdate(ss){
  ss = fetchSpreadsheet(ss);
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  const memberData = JSON.parse(docProps.getProperty('members')) || {};

  const totalMembers = memberData.memberOrder.length;
  let sheet, maxCols;

  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  if (config.pickemsInclude) {
    sheet = ss.getSheetByName('TOTAL');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
    overallMainFormulas(weeks,sheet,totalMembers,'TOT',true);

    sheet = ss.getSheetByName('RNK');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
    overallMainFormulas(weeks,sheet,totalMembers,'RNK',false);
  
    sheet = ss.getSheetByName('PCT');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
    overallMainFormulas(weeks,sheet,totalMembers,'PCT',true);
    
    if (!config.mnfExclude) {
      sheet = ss.getSheetByName('MNF');
      maxCols = sheet.getMaxColumns();
      overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
      overallMainFormulas(weeks,sheet,totalMembers,'MNF',true);
    }

    sheet = ss.getSheetByName('WINNERS');
    winnersFormulas(weeks,sheet);

    // Wkly Consensus now holds dollars written by updatePayouts(), not formulas -- left alone here
  }
}

// ============================================================================================================================================
// WEEKLY SHEETS
// ============================================================================================================================================

// WEEKLY Sheet Function - creates a sheet with provided week, members [array], and if data should be restored
function weeklySheet(ss,week,config,forms,memberData,displayEmpty,rebuild) {
  ss = ss || fetchSpreadsheet(ss);
  week = week || fetchWeek();
  let docProps = (!config || !forms || !memberData) ? PropertiesService.getDocumentProperties() : null;
  forms = forms || JSON.parse(docProps.getProperty('forms')) || {};

  if (!forms[week].gamePlan.pickemsInclude) {
    Logger.log(`⭕ Pick 'Ems not included in week ${week} form response, no weekly sheet needed`);
    ss.toast(`⭕ Pick 'Ems not included in week ${week} form response, no weekly sheet needed`);
    return null;
  }
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  
  // Insert Members
  let members = [];
  if (displayEmpty) { //|| forms[week].respondents == totalMembers) {
    members = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  } else {
    // Re-sorts based on memberOrder, then applies conversion to the name
    const sortedRespondents = forms[week].respondents.sort((a, b) => {
      const indexA = memberData.memberOrder.indexOf(a);
      const indexB = memberData.memberOrder.indexOf(b);
      const resolvedIndexA = indexA > -1 ? indexA : Infinity;
      const resolvedIndexB = indexB > -1 ? indexB : Infinity;
      return resolvedIndexA - resolvedIndexB;
    });
    members = sortedRespondents.map((id) => [memberData.members[id]?.name]);
  }
  let totalMembers = members.length;
  
  if (totalMembers <= 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(`⚠️ MEMBER ISSUE`, `There was an issue fetching the members to create the weekly sheet, make sure you've used the "Member Management" panel or waited for first submissions of the form before creating this sheet`,ui.ButtonSet.OK);
    Logger.log(`⚠️ Error fetching members to create weekly sheet`);
    return null;
  }

  const sheetName = weeklySheetPrefix + week;
  let sheet = ss.getSheetByName(sheetName);
  let existingData = null;

  // 1. (IF EXISTS) SCRAPE DATA: Perform a best-effort scrape of the existing sheet.
  if (sheet && rebuild) {
    Logger.log(`"${sheetName}" sheet exists. Attempting to preserve data before rebuild...`);
    ss.toast(`Attempting to gather and preserve any data that exists.`,`🔍 LOOKING FOR EXISTING DATA`)
    existingData = getExistingWeeklySheetData(ss, week, forms); // Use the new robust function
    if (existingData) {
      Logger.log(`💾 Data preservation successful. Scraped ${Object.keys(existingData.playerData).length} players.`);
      ss.toast(`Preserving existing data for week ${week}...`,`💾 SAVED EXISTING DATA`);
    } else {
      Logger.log(`🚫 Could not find valid data to preserve in sheet. It will be completely reset.`);
      ss.toast(`Could not find valid data to preserve in sheet. It will be completely reset.`,`🚫 NO OLD DATA FOUND`);
    }
  }

  // 2. NUKE & PAVE: Always start with a clean sheet.
  if (sheet) {
    sheet.clear();
    sheet.clearNotes();
    sheet.clearConditionalFormatRules();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
    Logger.log(`🧼 Sheet "${sheetName}" has been cleaned for rebuild.`);
    ss.toast(`Cleared out sheet for rebuilding and repopulating if possible`,`🧼 CLEANED ${weeklySheetPrefix}${week} SHEET`)
  } else {
    sheet = ss.insertSheet(sheetName, ss.getNumSheets() + 1);
    Logger.log(`⭐ Sheet "${sheetName}" created.`);
    ss.toast(`Created the new weekly sheet for populating with form data.`,`⭐ NEW ${weeklySheetPrefix}${week} SHEET`);
  }
  

  const contests = forms[week].gamePlan.games;
  const matchups = contests.length;
  const isAts = forms[week].gamePlan.pickemsAts;
  
  let diffCount = (totalMembers - 1) >= 5 ? 5 : (totalMembers - 1); // Number of results to display for most similar weekly picks (defaults to 5, or 1 fewer than the total member count, whichever is larger)

  const matchupRow = 1; // Row for all matchups
  const subHeaderRow = matchupRow + 1; // Row for denoting day of the week
  const entryRowStart = subHeaderRow + 1; // Row of first user input on weekly sheet
  const entryRowEnd = (entryRowStart - 1) + totalMembers; // Includes any header rows (entryRowStart-1) and adds two additional for final row of home/away splits and then bonus values
  const summaryRow = entryRowEnd + 1; // Row for group averages (away/home) and other calculated values
  const spreadRow = summaryRow + 1; // Recorded spreads (hidden if not ATS)
  const outcomeRow = summaryRow + 2; // Row for matchup outcomes
  const outcomeMarginRow = summaryRow + 3; // Row for margins
  const spreadOutcomeRow = summaryRow + 4; // Row for determining which team was the corret pick when including the spread
  const bonusRow = summaryRow + 5; // Row for adding bonus drop-downs
  const consensusRow = summaryRow + 6; // Pool 2: the crowd's majority pick for each matchup
  const rows = consensusRow; // Declare row variable, unnecessary, but easier to work with  
  const spreadToBonusRowCount = bonusRow - spreadRow; // Bottom area for use when highlighting for bonus presence
  let columns;
  
  // Adjust to the correct number of rows
  adjustRows(sheet,rows);
  
  let maxCols = sheet.getMaxColumns();
  
  // DATA GATHERING IF DATA RESTORE ACTIVE
  let commentCol, paidCol, tiebreakerCol = -1, tiebreaker2Col = -1;
  
  sheet.getRange(entryRowStart,1,totalMembers,1).setValues(members); 

  // Setting header values
  let headers = [`WEEK ${week}`,'⭐','🥇','💯','🤝','🎲','📊'];
  let subHeaders = [`${matchups} ${LEAGUE} Matchups`,'Picks','Rank','Percent','Consensus','Chances','']; // One blank for sparkline cell, will be merged
  let fontSizes = [18,16,16,16,16,16,16];
  let subFontSizes = [9,7,7,7,7,7,7];
  const subHeadersPriorLength = subHeaders.length;
  let bottomHeaders = ['Group Stats'];
  const pointsCol = subHeaders.indexOf('Picks') + 1;
  const rankCol = subHeaders.indexOf('Rank') + 1;
  const percentCol = subHeaders.indexOf('Percent') + 1;
  const consensusCol = subHeaders.indexOf('Consensus') + 1; // Pool 2: 1 = beat the crowd, 0 = did not
  const chancesCol = subHeaders.indexOf('Chances') + 1;
  const sparklinesCol = subHeaders.indexOf('Chances') + 2;

  sheet.getRange(summaryRow,1).setValue([bottomHeaders]);
  sheet.getRange(spreadRow,1).setValue('Spread Value');
  sheet.getRange(outcomeRow,1).setValue('Winner');
  sheet.getRange(outcomeMarginRow,1).setValue('Margin of Victory');
  sheet.getRange(spreadOutcomeRow,1).setValue('Winner Against the Spread');
  sheet.getRange(bonusRow,1).setValue('Bonus');
  sheet.getRange(consensusRow,1).setValue('Consensus pick');
  let widths = [130,50,50,50,50,50,50];

  
  // Setting headers for the week's matchups with format of 'AWAY' + '@' + 'HOME', then creating a data validation cell below each
  let firstMatchupCol = headers.length + 1;
  let mnfCol, winCol, spreads = [], subHeaderRowColors = [], bonuses = [], formatRules = [];
  let mnfCols = [];
  let newMatchupMap = {};
  let rule, matchupCol = 1;
  for ( let a in contests ) {
    let day = contests[a].dayName;
    let evening = contests[a].hour >= 17 ? true : false;
    let away = contests[a].awayTeam;
    let home = contests[a].homeTeam;
    newMatchupMap[`${away} @ ${home}`] = matchupCol++;
    // Establish start/stop of MNF games to record the tally
    if ( day == 'Monday' && evening ) {
      mnfCols.push(headers.length + 1);
      Logger.log(`🔍 MNF Added in Column ${headers.length + 1}: ${away}@${home}`);
    }
    let writeCell = sheet.getRange(subHeaderRow,firstMatchupCol+(matchups-1));
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=not(isblank(indirect("R${outcomeRow}C[0]",false)))`)
      .setBackground(dayColorsFilledObj[day] || '#b0b0b0')
      .setBold(true)
      .setRanges([writeCell]);
    rule.build();
    formatRules.push(rule);
    subHeaderRowColors.push(dayColorsObj[day] || '#e0e0e0');
    subHeaders.push(contests[a].dayName);
    const spread = contests[a].spread || '';
    spreads.push(spread);
    bonuses.push(contests[a].bonus);
    headers.push(`${away}\n@${home}`);
    widths.push(50);
    fontSizes.push(10);
    subFontSizes.push(7);
    rule = SpreadsheetApp.newDataValidation().requireValueInList([away,home,"TIE"], true).build();
    sheet.getRange(outcomeRow,headers.length).setDataValidation(rule);
  }
  if (existingData) {
    Logger.log(`🌏 Map created of new matchups with these values: ${JSON.stringify(newMatchupMap)}`);
  }

  const finalMatchupCol = headers.length;

  if (config.tiebreakerInclude) {
    Logger.log(`⚖️ Tiebreakers included in form; creating two tiebreaker columns and their difference columns`);
    // TIEBREAKER 1 -- combined score of the tiebreaker game
    headers.push('⚖️'); // Omitted if tiebreakers are removed
    subHeaders.push(`Tiebreaker 1`);
    widths.push(50);
    fontSizes.push(16);
    subFontSizes.push(7);
    tiebreakerCol = headers.length;
    headers.push('📏');
    subHeaders.push(`Difference 1`);
    widths.push(50);
    fontSizes.push(16);
    subFontSizes.push(7);
    sheet.getRange(matchupRow,tiebreakerCol).setNote(`Each member's guess at the combined score of the tiebreaker game (the late Monday night game, or the last game of a playoff week)`);
    sheet.getRange(matchupRow,tiebreakerCol+1).setNote(`The net difference between the combined score submitted and the actual combined score (in row ${outcomeRow}, column ${tiebreakerCol})`);

    // TIEBREAKER 2 -- winning team's score in the same game, only used if TB1 ties
    headers.push('🏁');
    subHeaders.push(`Tiebreaker 2`);
    widths.push(50);
    fontSizes.push(16);
    subFontSizes.push(7);
    tiebreaker2Col = headers.length;
    headers.push('📏');
    subHeaders.push(`Difference 2`);
    widths.push(50);
    fontSizes.push(16);
    subFontSizes.push(7);
    sheet.getRange(matchupRow,tiebreaker2Col).setNote(`Each member's guess at the WINNING team's score in the tiebreaker game, used only when Tiebreaker 1 is also tied`);
    sheet.getRange(matchupRow,tiebreaker2Col+1).setNote(`The net difference between the winning team score submitted and the actual winning score (in row ${outcomeRow}, column ${tiebreaker2Col})`);
  } else {
    Logger.log(`🚫 No tiebreakers included in form`);
  }

  headers.push('🏆');
  subHeaders.push(`Place`); // Finishing place, written by updatePayouts() once the week is scored
  widths.push(50);
  fontSizes.push(16);
  subFontSizes.push(7);
  winCol = headers.length;

  if (!config.mnfExclude && mnfCols.length > 0) {
    Logger.log(`🌙 MNF tracking included in pool; creating column`);
    headers.push('🌙');
    subHeaders.push('MNF'); // Added if user wants a MNF competition included
    widths.push(50);
    fontSizes.push(16);
    subFontSizes.push(7);
    mnfCol = headers.length;
    sheet.getRange(matchupRow,mnfCol).setNote(mnfCols.length > 1 ? `Displays number of correctly chosen MNF matchups of the possible ${mnfCols.length}` : `Displays whether the user correctly picked the MNF matchup`);
  } else {
    Logger.log(`🚫 No MNF tracking in the pool`);
  }

  if (!config.commentsExclude) {
    Logger.log(`💬 Comments allowed in the form; creating column`)
    headers.push('💬');
    subHeaders.push('Comments'); // Added to allow submissions to have amusing comments, if desired
    widths.push(150);
    fontSizes.push(16);
    subFontSizes.push(7);
    commentCol = headers.length;
    sheet.getRange(matchupRow,commentCol).setNote(`Column for member comments from the week ${week} form`);
  } else {
    Logger.log(`🤐 No comments allowed in form`);
  }
  
  // Wildcard column
  headers.push('🃏');
  subHeaders.push('Wildcard');
  widths.push(50);
  fontSizes.push(16);
  subFontSizes.push(7);
  const wildcardCol = headers.length;
  sheet.getRange(matchupRow,wildcardCol).setNote(`Represents the percent alignment to the median set of picks for week ${week}`);

  headers.push('🤝 COHESION');
  subHeaders.push('How many picks you differ from other members');
  const diffCol = headers.length;
  let finalCol = diffCol + (diffCount-1);
  headers.push(...Array(diffCount-1).fill(''));
  subHeaders.push(...Array(diffCount-1).fill(''));
  widths.push(...Array(diffCount).fill(90));
  fontSizes.push(...Array(diffCount).fill(10));
  subFontSizes.push(...Array(diffCount).fill(7));
  sheet.getRange(matchupRow,diffCol).setNote(`Displayed as the number of picks deviated from the next closests pickers`)
  
  // Add weekly checkboxes for payment status if configured
  const paidCheckboxes = config?.weeklyPaidTracking ? config.weeklyPaidTracking : false;
  if (paidCheckboxes) {
    Logger.log(`💵 Paid checkboxes included in sheet; creating column`);
    headers.push('💵');
    subHeaders.push(`Paid`);
    widths.push(70);
    fontSizes.push(16);
    subFontSizes.push(9);
    finalCol++;
    paidCol = finalCol;
    sheet.getRange(matchupRow,paidCol).setNote(`Tracking column for weekly payment status of members`);
  } else {
    Logger.log(`🚫 No paid checkboxes included in sheet`);
  }

  // Headers completed, now adjusting number of columns once headers are populated
  adjustColumns(sheet,finalCol);
  maxCols = sheet.getMaxColumns();

  sheet.getRange(matchupRow,1,1,headers.length).setValues([headers]);
  sheet.getRange(subHeaderRow,1,1,subHeaders.length).setValues([subHeaders]);

  // Note and subtitle setting
  sheet.getRange(matchupRow,pointsCol).setNote(config.bonusInclude ? `The number of correct points using bonus multipliers` : `The current amount of correct picks on the week`);
  sheet.getRange(matchupRow,rankCol).setNote(`Current weekly rank of each member`);
  sheet.getRange(matchupRow,consensusCol).setNote(`1 when a member beat the group's consensus picks for the week, 0 when they did not. Matching the consensus exactly does not count as beating it. Scored on raw correct picks -- the weekly bonus game does not count here. A game the NFL ends in a tie is a freebie worth +1 to every member and to the consensus. When members split 50/50 on a game the consensus has no pick (shown as SPLIT) and scores nothing there.`);
  sheet.getRange(matchupRow,percentCol).setNote(config.bonusInclude ? `Percent of picks correct (disregards bonus multipliers)` : `Percent of picks correct`);
  sheet.getRange(matchupRow,chancesCol).setNote(config.pickemsAts ? `Chance to finish with the most ${config.bonusInclude ? 'points':'correct picks'} on the week, accounts for spread probabilities${config.tiebreakerInclude ? ' but does not consider tiebreakers ' : ''}` : `Chance to finish with the most ${config.bonusInclude ? 'points':'correct picks'} on the week`);
  
  // Place spread values
  sheet.getRange(spreadRow,firstMatchupCol,1,spreads.length).setValues([spreads])
  
  // Set Data validation for margin
  sheet.getRange(outcomeMarginRow,firstMatchupCol,1,spreads.length).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());
  
  // Set Bonus values and validation
  let bonusRange = sheet.getRange(bonusRow,firstMatchupCol,1,bonuses.length);
  bonusRange.setValues([bonuses]);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(['1','2','3'],true).build();
  bonusRange.setDataValidation(rule);
  

  // Begin building functions
  // Create named ranges
  ss.setNamedRange(`${LEAGUE}_${week}`,sheet.getRange(matchupRow,firstMatchupCol,1,matchups)); // Then shortname versions of the matchups ( do have \n within )
  ss.setNamedRange(`${LEAGUE}_DAYS_${week}`,sheet.getRange(subHeaderRow,firstMatchupCol,1,matchups)); // Then shortname versions of the matchups ( do have \n within )
  ss.setNamedRange(`${LEAGUE}_BIAS_${week}`,sheet.getRange(summaryRow,firstMatchupCol,1,matchups)); // Then shortname versions of the matchups ( do have \n within )
  ss.setNamedRange(`${LEAGUE}_SPREADS_${week}`,sheet.getRange(spreadRow,firstMatchupCol,1,matchups)); // Spread values along bottom
  ss.setNamedRange(`${LEAGUE}_PICKEM_OUTCOMES_${week}`,sheet.getRange(outcomeRow,firstMatchupCol,1,matchups)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`,sheet.getRange(outcomeMarginRow,firstMatchupCol,1,matchups)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_ATS_OUTCOMES_${week}`,sheet.getRange(spreadOutcomeRow,firstMatchupCol,1,matchups)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_BONUS_${week}`,sheet.getRange(bonusRow,firstMatchupCol,1,matchups)); // Bonus multiplier for matchups
  ss.setNamedRange(`${LEAGUE}_PICKS_${week}`,sheet.getRange(entryRowStart,firstMatchupCol,totalMembers,matchups)); // All center data area (imported)

  if (config.tiebreakerInclude) {
    ss.setNamedRange(`${LEAGUE}_TIEBREAKER_${week}`,sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1));
    ss.setNamedRange(`${LEAGUE}_TIEBREAKER_${week}_OUTCOME`,sheet.getRange(outcomeRow,tiebreakerCol)); // Tiebreaker 1 Outcome
    let validRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0,150)
      .setHelpText('Must be an integer between 0 and 150')
      .build();
    sheet.getRange(outcomeRow,tiebreakerCol).setDataValidation(validRule);

    ss.setNamedRange(`${LEAGUE}_TIEBREAKER2_${week}`,sheet.getRange(entryRowStart,tiebreaker2Col,totalMembers,1));
    ss.setNamedRange(`${LEAGUE}_TIEBREAKER2_${week}_OUTCOME`,sheet.getRange(outcomeRow,tiebreaker2Col)); // Tiebreaker 2 Outcome
    let validRule2 = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0,100)
      .setHelpText(`Must be an integer between 0 and 100 (the winning team's score)`)
      .build();
    sheet.getRange(outcomeRow,tiebreaker2Col).setDataValidation(validRule2);
  }
  if (!config.commentsExclude) {
    ss.setNamedRange(`COMMENTS_${week}`,sheet.getRange(entryRowStart,commentCol,totalMembers,1));
  }

  // Pool 2 ranges: the per-member 1/0 flag column, and the crowd's pick for each matchup
  ss.setNamedRange(`${LEAGUE}_CONSENSUS_${week}`,sheet.getRange(entryRowStart,consensusCol,totalMembers,1));
  ss.setNamedRange(`${LEAGUE}_CONSENSUS_PICKS_${week}`,sheet.getRange(consensusRow,firstMatchupCol,1,matchups));

  const numPlayers = entryRowEnd - entryRowStart + 1;
  const effectiveOutcomeRow = isAts ? spreadOutcomeRow : outcomeRow;

  sheet.getRange(entryRowStart, 1, numPlayers, maxCols)
       .setBorder(null, null, true, null, false, true, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // These formulas are written for the FIRST row of the range (e.g., row 3).
  // Using relative R1C1 notation (like R[0]) allows them to automatically adjust for all other rows.

  // Define core range strings for reuse
  const picksRange = `R[0]C${firstMatchupCol}:R[0]C${finalMatchupCol}`;
  const allPicksRange = `R${entryRowStart}C${firstMatchupCol}:R${entryRowEnd}C${finalMatchupCol}`;
  const allWinnersRange = `R${entryRowStart}C${winCol}:R${entryRowEnd}C${winCol}`;
  const outcomesRange = `R${effectiveOutcomeRow}C${firstMatchupCol}:R${effectiveOutcomeRow}C${finalMatchupCol}`;
  const pointsCell = `R[0]C${pointsCol}`;
  const rankCell = `R[0]C${rankCol}`;
  const chancesCell = `R[0]C${chancesCol}`;
  const allChancesRange = `R${entryRowStart}C${chancesCol}:R${entryRowEnd}C${chancesCol}`;
  const allPointsRange = `R${entryRowStart}C${pointsCol}:R${entryRowEnd}C${pointsCol}`;
  const allSpreadsRange = `R${spreadRow}C${firstMatchupCol}:R${spreadRow}C${finalMatchupCol}`;
  const allBonusRange = `R${bonusRow}C${firstMatchupCol}:R${bonusRow}C${finalMatchupCol}`;

  // Points Formula (using efficient SUMPRODUCT)
  // A game the NFL ties is a freebie for everyone, so it is added here too -- otherwise this
  // column would read one lower than the points the weekly placement actually ranks on.
  const pointsFormula = `=IFERROR(IF(COUNTA(${outcomesRange}) > 0, SUMPRODUCT(--(${picksRange}=${outcomesRange}), ${allBonusRange}) + SUMPRODUCT(--(${outcomesRange}="TIE"), ${allBonusRange}),))`;

  // Rank Formula
  const rankFormula = `=IFERROR(IF(NOT(ISBLANK(${pointsCell})), RANK(${pointsCell}, ${allPointsRange}, 0),""))`;

  // Percent Correct Formula (using efficient SUMPRODUCT)
  const percentFormula = `=IFERROR(IF(COUNTA(${outcomesRange}) > 0, SUMPRODUCT(--(${picksRange}=${outcomesRange}), --(${outcomesRange}<>"")) / COUNTA(${outcomesRange}),""))`;

  // Chances formula (uses external function) for all rows
  const chancesFormula = `=calculateWinProbability(${allPicksRange},${outcomesRange},${allPointsRange},${allBonusRange},${allWinnersRange},${allSpreadsRange})`;

  // Sparkline Formula (leverages chances column adjacent)
  const sparklineFormula = `=IFERROR(IF(NOT(ISBLANK(${pointsCell})), SPARKLINE(MAX(${chancesCell},0.05),{"charttype","bar";"max",1;"color1",IF(${chancesCell}=max(${allChancesRange}),"#33ff7a",IF(${chancesCell}<(max(${allChancesRange})/3),"#ffa579","#ffe433"))}),),)`
  
  // Wildcard Formula (uses external function) for all rows
  const wildCardFormula = `=calculateWildcardScore(${allPicksRange})`;

  // Tiebreaker Difference Formulas (one per tiebreaker column, each reads the cell to its left)
  const tiebreakerDiffFormula = `=IFERROR(IF(OR(ISBLANK(R[0]C[-1]), ISBLANK(R${outcomeRow}C${tiebreakerCol})),, ABS(R[0]C[-1] - R${outcomeRow}C${tiebreakerCol})))`;
  const tiebreaker2DiffFormula = `=IFERROR(IF(OR(ISBLANK(R[0]C[-1]), ISBLANK(R${outcomeRow}C${tiebreaker2Col})),, ABS(R[0]C[-1] - R${outcomeRow}C${tiebreaker2Col})))`;

  // Similar Pickers Formula
  const similarPickersFormula = `=IFERROR(IF(ISBLANK(R[0]C${firstMatchupCol}),, TRANSPOSE(ARRAYFORMULA({(${matchups} - QUERY({R${entryRowStart}C1:R${entryRowEnd}C1, ARRAYFORMULA(MMULT(IF(${allPicksRange}=${picksRange},1,0),TRANSPOSE(ARRAYFORMULA(COLUMN(${picksRange})^0))))}, "select Col2 where Col1 <> '"&R[0]C1&"' order by Col2 desc, Col1 asc limit ${diffCount}")) & ": " & QUERY({R${entryRowStart}C1:R${entryRowEnd}C1, ARRAYFORMULA(MMULT(IF(${allPicksRange}=${picksRange},1,0),TRANSPOSE(ARRAYFORMULA(COLUMN(${picksRange})^0))))}, "select Col1 where Col1 <> '"&R[0]C1&"' order by Col2 desc, Col1 asc limit ${diffCount}")}))))`;

  // Apply formulas to ranges
  sheet.getRange(entryRowStart, pointsCol, numPlayers).setFormulaR1C1(pointsFormula);
  sheet.getRange(entryRowStart, rankCol, numPlayers).setFormulaR1C1(rankFormula);
  sheet.getRange(entryRowStart, percentCol, numPlayers).setFormulaR1C1(percentFormula);
  
  sheet.getRange(entryRowStart, chancesCol).setFormulaR1C1(chancesFormula); // Only in first cell--outputs an array
  sheet.getRange(entryRowStart, sparklinesCol, numPlayers).setFormulaR1C1(sparklineFormula);
  
  sheet.getRange(entryRowStart, wildcardCol).setFormulaR1C1(wildCardFormula);

  sheet.getRange(entryRowStart, diffCol, numPlayers).setFormulaR1C1(similarPickersFormula);

  // Apply conditional formulas
  if (config.tiebreakerInclude) {
    sheet.getRange(entryRowStart, tiebreakerCol + 1, numPlayers).setFormulaR1C1(tiebreakerDiffFormula);
    sheet.getRange(entryRowStart, tiebreaker2Col + 1, numPlayers).setFormulaR1C1(tiebreaker2DiffFormula);

    // WIN CASCADE: most correct -> closest combined score (TB1) -> closest winning
    // team score (TB2). Each level is skipped while its outcome cell is still blank,
    // and anyone who survives all three levels gets a 1, so unbroken ties produce
    // multiple winners and the weekly pot gets split.
    const weekComplete = `COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})=VALUE(REGEXEXTRACT(R${subHeaderRow}C1,"[0-9]+"))`;
    const tb1Outcome = `R${outcomeRow}C${tiebreakerCol}`;
    const tb2Outcome = `R${outcomeRow}C${tiebreaker2Col}`;
    const tb1Diffs = `R${entryRowStart}C${tiebreakerCol+1}:R${entryRowEnd}C${tiebreakerCol+1}`;
    const tb2Diffs = `R${entryRowStart}C${tiebreaker2Col+1}:R${entryRowEnd}C${tiebreaker2Col+1}`;
    const myDiff1 = `R[0]C${tiebreakerCol+1}`;
    const myDiff2 = `R[0]C${tiebreaker2Col+1}`;
    const bestPoints = `MAX(${allPointsRange})`;
    const minDiff1 = `MIN(FILTER(${tb1Diffs},${allPointsRange}=${bestPoints}))`;
    const minDiff2 = `MIN(FILTER(${tb2Diffs},(${allPointsRange}=${bestPoints})*(${tb1Diffs}=${minDiff1})))`;
    const atTopPoints = `${pointsCell}=${bestPoints}`;
    const passesTb1 = `IF(ISBLANK(${tb1Outcome}),TRUE,${myDiff1}=${minDiff1})`;
    const passesTb2 = `IF(OR(ISBLANK(${tb1Outcome}),ISBLANK(${tb2Outcome})),TRUE,${myDiff2}=${minDiff2})`;

    const tiebreakerWinnerFormula = `=IFERROR(IF(${weekComplete}, IF(AND(${atTopPoints},${passesTb1},${passesTb2}),1,0),),)`;
    sheet.getRange(allWinnersRange).setFormulaR1C1(tiebreakerWinnerFormula);

  } else {
    const noTiebreakerWinnerFormula = `=IFERROR(IF(COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})=VALUE(REGEXEXTRACT(R${subHeaderRow}C1,"[0-9]+")), IF(RANK(${pointsCell}, ${allPointsRange}, 0)=1, 1, 0)),)`;
    sheet.getRange(allWinnersRange).setFormulaR1C1(noTiebreakerWinnerFormula);
  }

  if (!config.mnfExclude && mnfCols.length > 0) {
    let mnfFormulaCore;

    if (mnfCols.length === 1) {
      // Logic for a single MNF game (more efficient)
      const singleCol = mnfCols[0];
      mnfFormulaCore = `-- (R[0]C${singleCol}=R${effectiveOutcomeRow}C${singleCol})`;
    } else {
      // Logic for multiple, non-contiguous MNF games
      // Creates array strings like "{R[0]C19, R[0]C20}"
      const picksArrayString = `{${mnfCols.map(c => `R[0]C${c}`).join(',')}}`;
      const resultsArrayString = `{${mnfCols.map(c => `R${effectiveOutcomeRow}C${c}`).join(',')}}`;
      mnfFormulaCore = `SUMPRODUCT(--(${picksArrayString}=${resultsArrayString}))`;
    }

    // Wrap the core logic in the standard IFERROR and readiness check
    const mnfFormula = `=IFERROR(IF(COUNTA(${outcomesRange}) > 0, ${mnfFormulaCore},""),"")`;
    sheet.getRange(entryRowStart, mnfCol, numPlayers).setFormulaR1C1(mnfFormula);
  }

  // Formula for the Home/Away split summary in the summary row
  const homeAwaySplitFormula = `=IFERROR(IF(COUNTA(R${entryRowStart}C[0]:R${entryRowEnd}C[0])=0,, LET(total_picks, COUNTA(R${entryRowStart}C[0]:R${entryRowEnd}C[0]), home_team, REGEXEXTRACT(R${matchupRow}C[0], "[A-Z]{2,3}$"), home_picks, COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0], home_team), away_team, REGEXEXTRACT(R${matchupRow}C[0], "^[A-Z]{2,3}"), IF(home_picks = total_picks/2, "SPLIT"&CHAR(10)&"50%", IF(home_picks > total_picks/2, home_team & CHAR(10) & ROUND(100*home_picks/total_picks,0)&"%", away_team & CHAR(10) & ROUND(100*(total_picks-home_picks)/total_picks,0)&"%")))))`;
  sheet.getRange(summaryRow, firstMatchupCol, 1, matchups).setFormulaR1C1(homeAwaySplitFormula);

  // ---- POOL 2: WEEKLY CONSENSUS ----------------------------------------------
  // The crowd's pick for each matchup: whichever team the majority took, or "TIE" on a 50/50.
  const consensusPickFormula = `=IFERROR(IF(COUNTA(R${entryRowStart}C[0]:R${entryRowEnd}C[0])=0,, LET(away, REGEXEXTRACT(R${matchupRow}C[0],"^[A-Z]{2,3}"), home, REGEXEXTRACT(R${matchupRow}C[0],"[A-Z]{2,3}$"), a, COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0], away), h, COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0], home), IF(a>h, away, IF(h>a, home, "SPLIT")))))`;
  sheet.getRange(consensusRow, firstMatchupCol, 1, matchups).setFormulaR1C1(consensusPickFormula);

  const consensusPicksRange = `R${consensusRow}C${firstMatchupCol}:R${consensusRow}C${finalMatchupCol}`;
  const outcomesRowRange = `R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol}`;
  // Two separate situations, per the pool rules:
  //   * A real NFL tie (outcome recorded as "TIE") is a freebie: nobody could have picked it, so
  //     every player AND the consensus each bank a point. It cannot change who beats the crowd,
  //     which is the point -- it neither helps nor hurts anyone.
  //   * A 50/50 pick split leaves the consensus with no pick for that game (shown as "SPLIT"),
  //     so the crowd simply scores nothing there while players score normally.
  const tieGameAdjustment = `SUMPRODUCT(--(${outcomesRowRange}="TIE"))`;
  const consensusRawCorrect = `SUMPRODUCT(--(${consensusPicksRange}=${outcomesRowRange}),--(${outcomesRowRange}<>""),--(${outcomesRowRange}<>"TIE"))`;

  // Crowd's adjusted correct count, shown in the Picks column of the consensus row
  sheet.getRange(consensusRow, pointsCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(${outcomesRowRange})=0,, ${consensusRawCorrect} + ${tieGameAdjustment}))`);
  // How many members beat the crowd — the divisor when the weekly consensus pot is split
  sheet.getRange(consensusRow, consensusCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(${outcomesRowRange})=0,, COUNTIF(R${entryRowStart}C${consensusCol}:R${entryRowEnd}C${consensusCol},1)))`);

  // Per-member flag. Raw correct is recomputed here rather than read from the Picks column so that
  // bonus multipliers, if they are ever switched on, cannot distort the consensus comparison.
  const playerRawCorrect = `SUMPRODUCT(--(R[0]C${firstMatchupCol}:R[0]C${finalMatchupCol}=${outcomesRowRange}),--(${outcomesRowRange}<>""),--(${outcomesRowRange}<>"TIE"))`;
  const weekSettled = `COUNTA(${outcomesRowRange})=VALUE(REGEXEXTRACT(R${subHeaderRow}C1,"[0-9]+"))`;
  const consensusFlagFormula = `=IFERROR(IF(${weekSettled}, IF(ISBLANK(R[0]C${firstMatchupCol}),, IF(${playerRawCorrect} + ${tieGameAdjustment} > R${consensusRow}C${pointsCol}, 1, 0)),),)`;
  sheet.getRange(entryRowStart, consensusCol, numPlayers).setFormulaR1C1(consensusFlagFormula);

  // Formula to calculate the winner based on the spread
  const spreadOutcomeFormula = `=IFERROR(IF(OR(ISBLANK(R${outcomeRow}C[0]), ISBLANK(R${outcomeMarginRow}C[0])),, LET(
    winner, R${outcomeRow}C[0],
    margin, R${outcomeMarginRow}C[0],
    spread_cell_text, R${spreadRow}C[0],
    full_matchup_text, R${matchupRow}C[0],
    
    favored_team, IFERROR(REGEXEXTRACT(spread_cell_text, "^[A-Z]{2,3}")),
    spread_line, IFERROR(VALUE(REGEXEXTRACT(spread_cell_text, "[-+][0-9\.]+"))),
    cover_number, ABS(spread_line),
    
    underdog_team, IFERROR(TRIM(SUBSTITUTE(SUBSTITUTE(full_matchup_text, favored_team,""), "@",""))),
    
    IF(winner = "TIE", underdog_team,
      IF(margin = cover_number, "TIE",
        IF(winner = favored_team,
          IF(margin > cover_number, favored_team, underdog_team),
          underdog_team
        )
      )
    )
  )))`;

  
  sheet.getRange(spreadOutcomeRow, firstMatchupCol, 1, matchups).setFormulaR1C1(spreadOutcomeFormula);
  
  // Points column headers and summary
  sheet.getRange(subHeaderRow, pointsCol).setFormulaR1C1(`=IF(COUNTIF(${allBonusRange},">1")>0, "Points", "Picks")`);
  sheet.getRange(summaryRow, pointsCol).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, "Average:"&CHAR(10)&TEXT(ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),"#.0"),),)`);

  // Rank column summary formula
  sheet.getRange(summaryRow, 3).setFormulaR1C1(`=IFERROR(IF(COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)>1,COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)&"-Way"&CHAR(10)&"Tie","Leader:"&CHAR(10)&INDEX(R${entryRowStart}C1:R${entryRowEnd}C1,MATCH(1,R${entryRowStart}C[0]:R${entryRowEnd}C[0],0),1)),)`); 

  // Percent summary formula
  sheet.getRange(summaryRow, 4).setFormulaR1C1(`=IFERROR(IF(COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})>2, AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),),)`);

  // Wildcard column summary formula
  sheet.getRange(summaryRow, wildcardCol).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),),)`);

  // Home/Away Bias summary formulas
  sheet.getRange(summaryRow, chancesCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(${allPicksRange})>10,"AWAY"&CHAR(10)&ROUND(100*(SUMPRODUCT(ARRAYFORMULA(--(REGEXEXTRACT(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol},"^[A-Z]{2,3}")=${allPicksRange}))))/COUNTA(${allPicksRange}),1)&"%","AWAY"),"AWAY")`);
  sheet.getRange(summaryRow, sparklinesCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(${allPicksRange})>10,"HOME"&CHAR(10)&ROUND(100*(SUMPRODUCT(ARRAYFORMULA(--(REGEXEXTRACT(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol},"[A-Z]{2,3}$")=${allPicksRange}))))/COUNTA(${allPicksRange}),1)&"%","HOME"),"HOME")`);
  
  // Tiebreaker and Winner columns
  if (config.tiebreakerInclude) {
    sheet.getRange(subHeaderRow, winCol).setFormulaR1C1(`=IF(COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)>1, "Tie", "Place")`);
    sheet.getRange(summaryRow, winCol).setFormulaR1C1(`=IFERROR(IF(NOT(ISBLANK(R${summaryRow}C${tiebreakerCol})), IF(COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)>1, COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)&"-WAY"&CHAR(10)&"TIE",),),)`);
    sheet.getRange(summaryRow, tiebreakerCol).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, "AVG"&CHAR(10)&ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),),)`);
    sheet.getRange(summaryRow, tiebreakerCol + 1).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, "AVG"&CHAR(10)&ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),),)`);
    sheet.getRange(summaryRow, tiebreaker2Col).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, "AVG"&CHAR(10)&ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),),)`);
    sheet.getRange(summaryRow, tiebreaker2Col + 1).setFormulaR1C1(`=IFERROR(IF(SUM(R${entryRowStart}C[0]:R${entryRowEnd}C[0])>0, "AVG"&CHAR(10)&ROUND(AVERAGE(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),1),),)`);
  } else {
    sheet.getRange(summaryRow, winCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})=VALUE(REGEXEXTRACT(R${subHeaderRow}C1,"[0-9]+")), IF(COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)>1, COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)&"-WAY"&CHAR(10)&"TIE", "DONE"),),)`);
    sheet.getRange(subHeaderRow, winCol).setFormulaR1C1(`=IFERROR(IF(COUNTA(R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol})=VALUE(REGEXEXTRACT(R${subHeaderRow}C1,"[0-9]+")), IF(COUNTIF(R${entryRowStart}C[0]:R${entryRowEnd}C[0],1)=0, "Tie", "Win"), "Win"),)`);
  }
  
  // MNF Summary Logic
  if (mnfCols && mnfCols.length > 0) {
    // Dynamically create a SUM of SUMPRODUCTs for each MNF column
    const correctPicksSumString = mnfCols.map(col => 
      `SUMPRODUCT(--(R${entryRowStart}C${col}:R${entryRowEnd}C${col}=R${effectiveOutcomeRow}C${col}))`
    ).join('+');
    
    const totalPicks = totalMembers * mnfCols.length;
    
    const mnfSummaryFormula = `=IFERROR(IF(AND(COUNTIF({${mnfCols.map(c => `R${effectiveOutcomeRow}C${c}`).join(',')}},"<>")=${mnfCols.length}), "MNF" & CHAR(10) & ROUND(100*(${correctPicksSumString})/${totalPicks},1)&"%",),)`;
    
    sheet.getRange(summaryRow, mnfCol).setFormulaR1C1(mnfSummaryFormula);
  }

  // Similar pickers for whole group
  sheet.getRange(summaryRow,diffCol).setFormulaR1C1(`=iferror(if(ISBLANK(R[0]C${firstMatchupCol}),,transpose(query({arrayformula((counta(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol})-mmult(arrayformula(if(R${entryRowStart}C${firstMatchupCol}:R${entryRowEnd}C${finalMatchupCol}=arrayformula(regexextract(R${totalMembers+3}C${firstMatchupCol}:R${totalMembers+3}C${finalMatchupCol},"[A-Z]+")),1,0)),transpose(arrayformula(if(arrayformula(len(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol}))>1,1,1)))))&": "&R${entryRowStart}C1:R${entryRowEnd}C1),mmult(arrayformula(if(R${entryRowStart}C${firstMatchupCol}:R${entryRowEnd}C${finalMatchupCol}=arrayformula(regexextract(R${totalMembers+3}C${firstMatchupCol}:R${totalMembers+3}C${finalMatchupCol},"[A-Z]+")),1,0)),transpose(arrayformula(if(arrayformula(len(R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol}))>1,1,1))))},"select Col1 order by Col2 desc, Col1 desc limit ${diffCount}"))))`);

  // --- Configuration for Conditional Formatting ---
  const bonusCount = 3;
  // Using an object for parity makes the code more self-documenting
  const parities = {
    even: { fn: 'iseven' },
    odd:  { fn: 'isodd'  }
  };

  // --- Define the base formulas ONCE using supported functions ---
  const outcomeRowRef = `INDIRECT("R${isAts ? spreadOutcomeRow : outcomeRow}C[0]", FALSE)`;
  const thisCellRef = `INDIRECT("R[0]C[0]", FALSE)`;
  const thisRowRef = `INDIRECT("R[0]C1", FALSE)`;
  const matchupRef = `INDIRECT("R${matchupRow}C[0]", FALSE)`;
  const bonusRef = `INDIRECT("R${bonusRow}C[0]", FALSE)`;

  const baseFormulas = {
    // Correct picks are the highest priority. Checks if the pick matches the outcome.
    correct: `AND(${outcomeRowRef}=${thisCellRef}, NOT(ISBLANK(${outcomeRowRef})))`,
    
    // Incorrect picks are next. This is any cell in a completed game that isn't correct.
    incorrect: `AND(${outcomeRowRef}<>${thisCellRef}, NOT(ISBLANK(${outcomeRowRef})), NOT(ISBLANK(${thisCellRef})))`,
    
    // Home picks are for games not yet played.
    // Formula: this cell's value = the text to the RIGHT of the "@" in the matchup row.
    home: `AND(ISBLANK(${outcomeRowRef}), NOT(ISBLANK(${thisCellRef})), ${thisCellRef}=TRIM(RIGHT(${matchupRef}, LEN(${matchupRef})-FIND("@",${matchupRef}))))`,
    
    // Away picks are for games not yet played.
    // Formula: this cell's value = the text to the LEFT of the "@" in the matchup row.
    away: `AND(ISBLANK(${outcomeRowRef}), NOT(ISBLANK(${thisCellRef})), ${thisCellRef}=TRIM(LEFT(${matchupRef}, FIND("@",${matchupRef})-1)))`
  };

  // --- Programmatically Generate and Apply Formatting Rules ---
  range = sheet.getRange(allPicksRange);

  // Uses top-level pickColors object
  // Rules are evaluated from top to bottom. `correct` and `incorrect` must come first.

  // Evaluate from top to bottom: Correct -> Incorrect -> Home -> Away
  for (const [type, cfg] of Object.entries(pickColors)) {
    for (const parity of Object.values(parities)) {
      const startColor = parity.fn === 'iseven' ? cfg.even : cfg.odd;
      const gradient = typeof hexGradient === 'function'
        ? hexGradient(startColor, cfg.end, bonusCount)
        : [startColor];

      // Loop backwards so higher bonus tiers evaluate first
      for (let i = gradient.length - 1; i >= 0; i--) {
        const bonusLevel = i + 1;
        const baseFormula = baseFormulas[type];
        const parityCondition = `${parity.fn}(ROW(${thisRowRef}))`;

        let finalFormula = `=AND(${baseFormula}, ${parityCondition}`;
        if (i > 0) {
          finalFormula += `, ${bonusRef}=${bonusLevel}`;
        }
        finalFormula += `)`;

        const ruleBuilder = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(finalFormula)
          .setBackground(gradient[i])
          .setRanges([range]);

        if (cfg.font) {
          ruleBuilder.setFontColor(cfg.font);
        }

        formatRules.push(ruleBuilder.build());
      }
    }
  }

  // NAMES COLUMN NAMED RANGE
  range = sheet.getRange(entryRowStart,1,totalMembers,1);
  ss.setNamedRange(`NAMES_${week}`,range);

  // TOTALS GRADIENT RULE
  range = sheet.getRange(entryRowStart,pointsCol,totalMembers,1);
  ss.setNamedRange(`TOT_${week}`,range);
  let formatRuleTotals = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#75F0A1')
    .setGradientMinpoint('#FFFFFF')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleTotals);
  // RANKS GRADIENT RULE
  range = sheet.getRange(entryRowStart,rankCol,totalMembers,1);
  ss.setNamedRange(`RNK_${week}`,range);
  let formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#FF9B69', SpreadsheetApp.InterpolationType.NUMBER, members.length)
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, members.length/2)
    .setGradientMinpointWithValue('#5EDCFF', SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleRanks);
  // PERCENT GRADIENT RULE
  ss.setNamedRange(`PCT_${week}`,sheet.getRange(entryRowStart,percentCol,totalMembers,1)); // Range for formatting below includes summary row
  sheet.getRange(entryRowStart,percentCol,totalMembers+1,1).setNumberFormat('##0.0%');
  formatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#75F0A1', SpreadsheetApp.InterpolationType.NUMBER, '.70')
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '.60')
    .setGradientMinpointWithValue('#FF9B69', SpreadsheetApp.InterpolationType.NUMBER, '.50')
    .setRanges([sheet.getRange(entryRowStart,percentCol,totalMembers+1,1)])
    .build());
  
  // CHANCES GRADIENT RULE  '#33ff7a',IF(E3<0.33,'#ffa579','#ffe433')
  range =  sheet.getRange(entryRowStart,chancesCol,totalMembers,1);
  ss.setNamedRange(`CHANCES_${week}`,range);
  range.setNumberFormat('##0.0%');
  formatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#33ff7a')
    .setGradientMidpointWithValue('#ffe433', SpreadsheetApp.InterpolationType.PERCENT, '50')
    .setGradientMinpoint('#ffa579')
    .setRanges([range])
    .build());

  // WILDCARD GRADIENT RULE  '#33ff7a',IF(E3<0.33,'#ffa579','#ffe433')
  ss.setNamedRange(`WILDCARD_${week}`,sheet.getRange(entryRowStart,wildcardCol,totalMembers,1)); // Range for formatting below includes summary row
  range.setNumberFormat('##0.0%');
  let formatRuleWildcard = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#fca503', SpreadsheetApp.InterpolationType.NUMBER, '0.50')
    .setGradientMidpointWithValue('#ffe433', SpreadsheetApp.InterpolationType.NUMBER, '0.25')
    .setGradientMinpointWithValue('#7dfffb', SpreadsheetApp.InterpolationType.NUMBER, '0.00')
    .setRanges([sheet.getRange(entryRowStart,wildcardCol,totalMembers+1,1)])
    .build();
  formatRules.push(formatRuleWildcard);

  // WINNER COLUMN RULE
  range = sheet.getRange(entryRowStart,winCol,totalMembers,1);
  ss.setNamedRange(`WIN_${week}`,range);
  // Paid places are shaded from strongest to faintest; anything outside the money stays plain.
  // "Place" is written by updatePayouts() when a week closes, so 1st here matches Wkly Payout.
  const placeColors = hexGradient('#75F0A1','#FFFFFF',8);
  for (let place = 1; place <= 7; place++) {
    let placeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(place)
      .setBackground(placeColors[place - 1])
      .setBold(place === 1)
      .setRanges([range])
      .build();
    formatRules.push(placeRule);
  }
  // WINNER NAME RULE
  range = sheet.getRange(entryRowStart,winCol,totalMembers,1);
  let formatRuleWinnerName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=indirect("R[0]C${winCol}",false)=1`)
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleWinnerName);

  // MNF GRADIENT RULE
  if (!config.mnfExclude && mnfCols.length > 0) {
    range = sheet.getRange(entryRowStart,mnfCol,totalMembers,1);
    ss.setNamedRange(`MNF_${week}`,range);
    // formatRuleMNFEmpty = SpreadsheetApp.newConditionalFormatRule()
    //   .whenCellEmpty()
    //   .setFontColor('#FFFFFF')
    //   .setBackground('#FFFFFF')
    //   .setRanges([range])
    //   .build();
    // formatRules.push(formatRuleMNFEmpty);
    formatRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(1)
      .setFontColor('#FFFFFF')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build());
    if (mnfCols.length > 1) { // Rules for when there are multiple MNF games
      formatRules.push(SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint("#FFF624") // Max value of all correct picks, min 1
        .setGradientMinpoint("#FFFFFF") // Min value of all correct picks  
        .setRanges([range])
        .build());
    } else { // Rules for single MNF game 
      formatRules.push(SpreadsheetApp.newConditionalFormatRule()
        .setBackground("#FFF624")
        .setFontColor("#FFF624")
        .whenNumberEqualTo(1)
        .setRanges([range])
        .build());
    }
  }

  // DIFFERENCE TIEBREAKER COLUMN FORMATTING
  if (config.tiebreakerInclude) {
    // Both tiebreaker columns get the same treatment
    for (const tbCol of (tiebreaker2Col > 0 ? [tiebreakerCol, tiebreaker2Col] : [tiebreakerCol])) {
      let offsets = [1,3,5,10,15,20,20];
      let offsetColors = hexGradient('#33FF7A','#FFFFFF',offsets.length);
      for (let a = 0; a < offsets.length; a++) {
        let rule;
        if (a < (offsets.length - 1)) {
          rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[0]",false))),abs(indirect("R[0]C[0]",false)-indirect("R${outcomeRow}C[0]:R${outcomeRow}C[0]",false))<=${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(entryRowStart,tbCol,totalMembers,1)])
            .build();
        } else {
          rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[0]",false))),abs(indirect("R[0]C[0]",false)-indirect("R${outcomeRow}C[0]:R${outcomeRow}C[0]",false))>${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(entryRowStart,tbCol,totalMembers,1)])
            .build();        
        }
        formatRules.push(rule);
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[0]",false))),abs(value(regexextract(indirect("R[0]C[0]",false),"[0-9]+"))-indirect("R${outcomeRow}C[0]:R${outcomeRow}C[0]",false))<=${offsets[a]},)`)
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(summaryRow,tbCol)])
          .build();
        formatRules.push(rule);
      }
      offsetColors = hexGradient('#FFFFFF','#666666',offsets.length);
      for (let a = 0; a < offsets.length; a++) {
        let rule;
        let ruleOffsets;
        if (a < (offsets.length - 1)) {
          rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[-1]",false))),indirect("R[0]C[0]",false)<=${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(entryRowStart,tbCol+1,totalMembers,1)])
            .build();
          ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[-1]",false))),value(regexextract(indirect("R[0]C[0]",false),"[0-9]+"))<=${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(summaryRow,tbCol+1)])
            .build();
        } else {
          rule = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[-1]",false))),indirect("R[0]C[0]",false)>${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(entryRowStart,tbCol+1,totalMembers,1)])
            .build();
          ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(`=if(not(isblank(indirect("R${outcomeRow}C[-1]",false))),value(regexextract(indirect("R[0]C[0]",false),"[0-9]+"))>${offsets[a]},)`)
            .setBackground(offsetColors[a])
            .setRanges([sheet.getRange(summaryRow,tbCol+1)])
            .build();              
        }
        formatRules.push(rule);
        formatRules.push(ruleOffsets);
      }
      // ADD ADDITIONAL COLOR VARIATION BASED ON TIEBREAKER VALUE PRESENT HERE
      let formatRuleTiebreakerEmptyAndDone = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=and(isblank(indirect("R[0]C[0]",false)),counta(indirect("R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol}",false))>=columns(indirect("R${outcomeRow}C${firstMatchupCol}:R${outcomeRow}C${finalMatchupCol}",false)))`)
        .setBackground("#FF3FC7")
        .setRanges([sheet.getRange(outcomeRow,tbCol)])
        .build();
      formatRules.push(formatRuleTiebreakerEmptyAndDone);
      let formatRuleTiebreakerEmpty = SpreadsheetApp.newConditionalFormatRule()
        .whenCellEmpty()
        .setBackground("#CCCCCC")
        .setRanges([sheet.getRange(outcomeRow,tbCol)])
        .build();
      formatRules.push(formatRuleTiebreakerEmpty);
      range = sheet.getRange(entryRowStart,tbCol,totalMembers,1);
      let formatRuleDiff = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint("#B7B7B7")
        .setGradientMinpoint("#FFFFFF")
        .setRanges([range])
        .build();
      formatRules.push(formatRuleDiff);
    }
  }

  // POOL 2 CONSENSUS COLUMN FORMATTING
  range = sheet.getRange(entryRowStart,consensusCol,totalMembers,1);
  let formatRuleBeatConsensus = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleBeatConsensus);
  let formatRuleMissedConsensus = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground('#FFFFFF')
    .setFontColor('#CCCCCC')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleMissedConsensus);

  // PREFERENCE COLOR SCHEMES
  let homeAwayPercents = [90,80,70,60,50];
  let awayColors = ['#FFFB7D','#FFFC96','#FFFCB0','#FFFDC9','#FFFEE3'];
  let homeColors = ['#7DFFFB','#96FFFC','#B0FFFC','#C9FFFD','#E3FFFE'];
  let awayFormula = `=and(regexextract(indirect("R[0]C[0]",false),"[A-Z]{2,3}")=regexextract(indirect("R${matchupRow}C[0]",false),"[A-Z]{2,3}"),value(regexextract(indirect("R[0]C[0]",false),"[0-9\.]+"))>=%%)`; // Replaceable "%%" for inserting percent number
  let homeFormula = `=and(regexextract(indirect("R[0]C[0]",false),"[A-Z]{2,3}")=regexextract(right(indirect("R${matchupRow}C[0]",false),3),"[A-Z]{2,3}"),value(regexextract(indirect(\"R[0]C[0]",false),"[0-9\.]+"))>=%%)`; // Replaceable "%%" for inserting percent number
  range = sheet.getRange(summaryRow,firstMatchupCol,1,matchups); // Summary row of matchups
  for (let a = 0; a < homeAwayPercents.length; a++) {
    let formula = awayFormula.replace('%%',homeAwayPercents[a]);

    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(awayColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);

    formula = homeFormula.replace('%%',homeAwayPercents[a]);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(homeColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);    
  }

  // MATCHUP WEIGHTING RULE
  let formatRuleWeightedThree, formatRuleWeightedTwo;
  const topBonusHighlightRange = sheet.getRange(matchupRow,firstMatchupCol,1,matchups); // Top Bar of matchups
  const bottomBonusHighlightRange = sheet.getRange(spreadRow,firstMatchupCol,spreadToBonusRowCount,matchups); // final rows of Spread, Winner, Margin, ATS Winner, & Bonus
  formatRuleWeightedThree = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=and(not(isblank(indirect("R[0]C[0]",false))),or(and(indirect("R${bonusRow}C[0]",false)=2,countif(indirect("R${bonusRow}C${firstMatchupCol}:R${bonusRow}C${finalMatchupCol}",false),3)=0),indirect("R${bonusRow}C[0]",false)=3))`)
    .setBackground('#9C9C97')
    .setRanges([topBonusHighlightRange, bottomBonusHighlightRange]) // LEGACY VERSION sheet.getRange(`R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol}`),sheet.getRange(`R${spreadRow}C${firstMatchupCol}:R${bonusRow}C${finalMatchupCol}`)])
    .build();
  formatRules.push(formatRuleWeightedThree);
  formatRuleWeightedTwo = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=and(not(isblank(indirect("R[0]C[0]",false))),indirect("R${bonusRow}C[0]",false)=2)`)
    .setBackground('#949376')
    .setRanges([topBonusHighlightRange, bottomBonusHighlightRange]) // LEGACY VERSION sheet.getRange(`R${matchupRow}C${firstMatchupCol}:R${matchupRow}C${finalMatchupCol}`),sheet.getRange(`R${spreadRow}C${firstMatchupCol}:R${bonusRow}C${finalMatchupCol}`)])
    .build();
  formatRules.push(formatRuleWeightedTwo);
  
  // Format rules for difference columns to emphasize the most common picker
  let commonPickersGradient = hexGradient('#46f081','#e4f0e8',8);
  const commonPickersFormula = `=value(regexextract(indirect("R[0]C[0]",false),"[0-9]+"))=`;
  range = sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount);
  for (let a = 0; a < commonPickersGradient.length; a++) {
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`${commonPickersFormula}${a}`)
      .setBackground(commonPickersGradient[a])
      .setRanges([range])
      .build();
    formatRules.push(rule);
  }
  
  // Add conditional formatting rules to indicate paid status (added last to take lowest priority)
  if (paidCheckboxes) {
    sheet.getRange(entryRowStart,paidCol,totalMembers,1).insertCheckboxes().setFontSize(11).setHorizontalAlignment('center');
    let finalPaidCell = sheet.getRange(entryRowEnd+1,paidCol);
    finalPaidCell.setHorizontalAlignment('center')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setFontWeight('bold')
      .setNumberFormat("##.#%")
      .setFormulaR1C1(`=if(and(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),"ALL PAID",if(arrayformula(and(NOT(R${entryRowStart}C[0]:R${entryRowEnd}C[0]))),"UNPAID",round(countif(R${entryRowStart}C[0]:R${entryRowEnd}C[0],true)/counta(R${entryRowStart}C[0]:R${entryRowEnd}C[0]),3)))`);
    let formatRuleAllPaid = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ALL PAID')
      .setBackground('#b5fff2')
      .setRanges([finalPaidCell])
      .build();
    formatRules.push(formatRuleAllPaid);
    let formatRuleAllNotPaid = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('UNPAID')
      .setBackground('#ffd5b5')
      .setRanges([finalPaidCell])
      .build();
    formatRules.push(formatRuleAllNotPaid);
    let formatRulePartialPaid = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#b5fff2", SpreadsheetApp.InterpolationType.NUMBER, ".99")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".50")
      .setGradientMinpointWithValue("#ffd5b5", SpreadsheetApp.InterpolationType.NUMBER, ".01")
      .setRanges([finalPaidCell])
      .build();
    formatRules.push(formatRulePartialPaid);
    sheet.setColumnWidth(paidCol,70);
    const nameBlockRange = sheet.getRange(entryRowStart,1,totalMembers,firstMatchupCol-1);
    const paidColRange = sheet.getRange(entryRowStart,paidCol,totalMembers,1);
    let formatRulePaid = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=indirect("R[0]C${paidCol}",false)=true`)
      .setBackground('#f0fffc')
      .setRanges([nameBlockRange,paidColRange])
      .build();      
    formatRules.push(formatRulePaid);
    let formatRuleUnpaid = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=indirect("R[0]C${paidCol}",false)=false`)
      .setBackground('#fff3eb')
      .setItalic(true)
      .setRanges([nameBlockRange,paidColRange])
      .build();
    formatRules.push(formatRuleUnpaid);
  }
 
  // Sets all formerly pushed rules to the sheet
  sheet.setConditionalFormatRules(formatRules);

  // Setting size, alignment, frozen columns
  columns = sheet.getMaxColumns();
  sheet.getRange(1,1,rows,columns)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontSize(10)
    .setFontFamily("Montserrat");
  sheet.getRange(subHeaderRow,2,1,subHeadersPriorLength - 1)
    .setFontSize(8)
    .setFontWeight('bold');
  sheet.getRange(subHeaderRow,subHeaders.indexOf('Chances')+1,1,2).mergeAcross();

  sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount).setHorizontalAlignment('left');
  if (!config.commentsExclude) {
    sheet.getRange(2,commentCol,totalMembers+1,1).setHorizontalAlignment('left');
  }

  sheet.getRange(1,1,summaryRow,1)
    .setHorizontalAlignment('left');
 
  sheet.setFrozenColumns(firstMatchupCol-1);
  sheet.setFrozenRows(subHeaderRow);
  sheet.getRange(1,1,1,columns)
    .setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet.setRowHeights(1,rows,21);

  sheet.getRange(matchupRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.setRowHeight(matchupRow,50);
  sheet.getRange(matchupRow,1).setHorizontalAlignment('center');
  
  sheet.getRange(subHeaderRow,firstMatchupCol,1,matchups).setFontSize(7);
  sheet.getRange(subHeaderRow,1,1,maxCols).setBackground('#CCCCCC');
  sheet.getRange(subHeaderRow,firstMatchupCol,1,subHeaderRowColors.length).setBackgrounds([subHeaderRowColors]);
  sheet.getRange(subHeaderRow,1).setHorizontalAlignment('left');
  
  // Lower area black background, white text, and font size 10
  sheet.getRange(spreadRow,1,bonusRow-spreadRow+1,maxCols)
    .setHorizontalAlignment('center')
    .setBackground('black')
    .setFontColor('white')
    .setFontSize(10);
  // Spread row to bonus row formatting
  sheet.getRange(spreadRow,1,consensusRow-spreadRow+1,firstMatchupCol-1)
    .mergeAcross().setHorizontalAlignment('right');
  // Smaller spread values to fit widths
  sheet.getRange(spreadRow,1,1,maxCols).setFontSize(8);
  // Bold on these 
  sheet.getRange(outcomeRow,1,1,maxCols).setFontWeight('bold');   
  sheet.getRange(spreadOutcomeRow,1,1,maxCols).setFontWeight('bold');

  if (!config.bonusInclude) {
    sheet.hideRows(bonusRow);
  }

  // Consensus row reads as a summary line rather than a data row
  sheet.getRange(consensusRow,1,1,maxCols).setBackground('#EDE7F6').setFontWeight('bold');
  sheet.getRange(consensusRow,1).setNote(`The group's majority pick for each matchup, or "SPLIT" when members divide 50/50. The Picks column shows the consensus score for the week (raw correct picks, plus a freebie for any game the NFL ties); the Consensus column shows how many members beat it.`);

  if (!isAts) {
    sheet.hideRows(spreadRow);
    sheet.hideRows(outcomeMarginRow);
    sheet.hideRows(spreadOutcomeRow);
  }

  sheet.setRowHeight(summaryRow,40);
  sheet.getRange(summaryRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.getRange(summaryRow,1,1,maxCols-diffCount).setBackground('#CCCCCC');
  sheet.getRange(summaryRow,chancesCol).setBackground(awayColors[1]);
  sheet.getRange(summaryRow,sparklinesCol).setBackground(homeColors[1]);

  // GROUP AVG POINTS/PICKS
  sheet.getRange(summaryRow,2).setFontSize(8).setBackground('#75F0A1');
  // GROUP LEADER/TIE
  sheet.getRange(summaryRow,3).setFontSize(8).setBackground('#5EDCFF');
  
  // MERGE Different picker columns
  sheet.getRange(1,diffCol,2,diffCount)
    .setHorizontalAlignment('left')
    .mergeAcross();

  if (config.tiebreakerInclude) {
    sheet.getRange(outcomeRow,tiebreakerCol).setNote('Enter the COMBINED score of the tiebreaker game (late Monday night game, or the last game of a playoff week) to complete the week and designate a winner');
    sheet.getRange(outcomeRow,tiebreaker2Col).setNote(`Enter the WINNING team's score in that same tiebreaker game — only used when Tiebreaker 1 is also tied`);
  }

  let lastWidthsValue, lastHeaderFontValue, lastSubHeaderFontValue;
  let widthsHeldCount = 0, headerHeldCount = 0, subHeaderHeldCount = 0;
  let widthsStartCol = 1, headerStartCol = 1, subHeaderStartCol = 1;

  // Set all cell/column specific sizes and formats
  for (let a = 0; a <= widths.length; a++) {
    // Handle Column Widths
    if (a === 0) {
      // First iteration - initialize
      lastWidthsValue = widths[a];
      widthsHeldCount = 1;
      widthsStartCol = 1;
    } else if (a === widths.length || lastWidthsValue !== widths[a]) {
      // Apply the batch when value changes or at end
      sheet.setColumnWidths(widthsStartCol, widthsHeldCount, lastWidthsValue);
      if (a < widths.length) {
        // Start new batch
        lastWidthsValue = widths[a];
        widthsHeldCount = 1;
        widthsStartCol = a + 1;
      }
    } else {
      // Same value, continue batch
      widthsHeldCount++;
    }
    
    // Handle Header Font Sizes (Row 1)
    if (a === 0) {
      // First iteration - initialize
      lastHeaderFontValue = fontSizes[a];
      headerHeldCount = 1;
      headerStartCol = 1;
    } else if (a === fontSizes.length || lastHeaderFontValue !== fontSizes[a]) {
      // Apply the batch when value changes or at end
      sheet.getRange(1, headerStartCol, 1, headerHeldCount).setFontSize(lastHeaderFontValue);
      if (a < fontSizes.length) {
        // Start new batch
        lastHeaderFontValue = fontSizes[a];
        headerHeldCount = 1;
        headerStartCol = a + 1;
      }
    } else {
      // Same value, continue batch
      headerHeldCount++;
    }
    
    // Handle SubHeader Font Sizes (Row 2)
    if (a === 0) {
      // First iteration - initialize
      lastSubHeaderFontValue = subFontSizes[a];
      subHeaderHeldCount = 1;
      subHeaderStartCol = 1;
    } else if (a === subFontSizes.length || lastSubHeaderFontValue !== subFontSizes[a]) {
      // Apply the batch when value changes or at end
      sheet.getRange(2, subHeaderStartCol, 1, subHeaderHeldCount).setFontSize(lastSubHeaderFontValue);
      
      if (a < subFontSizes.length) {
        // Start new batch
        lastSubHeaderFontValue = subFontSizes[a];
        subHeaderHeldCount = 1;
        subHeaderStartCol = a + 1;
      }
    } else {
      // Same value, continue batch
      subHeaderHeldCount++;
    }
  }

  // RESTORE STATE: If we successfully scraped data, repopulate it now.
  if (rebuild && existingData && newMatchupMap) {
    Logger.log(`🔄 Restoring preserved data into new sheet structure...`);
    remapAndRepopulateData(ss, week, existingData, newMatchupMap, members.map(m => m[0]));
    ss.toast(`✅ Data successfully restored for week ${week}.`, 'SUCCESS');
  } else if (existingData && !newMatchupMap) {
      Logger.log(`⚠️ ERROR: Scraped old data but failed to get a new game map. Data could not be restored.`);
      ss.toast(`ERROR: Could not restore data.`, `⚠️ ERROR`);
  }

  // Adds tab colors
  weeklySheetTabColors(ss,week,rebuild && existingData && newMatchupMap); 

  const text = `Completed creation of pick 'ems week ${week} sheet.`;
  Logger.log(`✅ $(text)`);
  ss.toast(text,`✅ SUCCESS`);
  return sheet;
}

// WEEKLY SHEET COLORATION - Adds a color to the weekly tabs that exist and uses the "dayColorsFilled" array [global variable]
function weeklySheetTabColors(ss, week, all) {
  let maxWeek = 1, old = false;
  if (all) {
    maxWeek = weeklySheetFindLargest(ss);
  }
  week = week || fetchWeek();
  if (week < maxWeek) old = true;
  week = week < maxWeek ? maxWeek : week;
  let icon = week <= 10 ? numberMap[week] : (numberMap[Math.floor(week/10)] + numberMap[week - Math.floor(week/10)*10]);
  if (old) {
    Logger.log(`❕ Detected regression in weekly rebuild, color coding based on max week sheet of ${icon} in the spreadsheet`);
  } else {
    Logger.log(`🔄 Applying custom coloration based on the start week of ${week}`)
  }
  ss = ss || fetchSpreadsheet();
  let sheet = ss.getSheetByName(`${weeklySheetPrefix}${week}`);
  try {
    if (!sheet) {
      throw new Error();
    }
    let colors = [...dayColorsFilled];
    colors.push(winnersTabColor); // Adds a bright yellow to the end of the array for the active week tab
    let week = parseInt(sheet.getName().replace(weeklySheetPrefix,''));
    sheet.setTabColor(colors[colors.length-1]);
    colors.pop();
    for (let a = (week - 1); a > 0; a--) {
      let sheet = ss.getSheetByName(weeklySheetPrefix + a);
      if (sheet) {
        sheet.setTabColor(colors[colors.length-1]);
      }
      if (colors.length > 1) {
        colors.pop();
      }
    }
    Logger.log(`🎨 Changed all colors of tabs to reflect week shift to week ${week}`);
  }
  catch (err) {
    Logger.log(`❗ Error assigning colors to weekly sheet tabs | ERROR: ${err.stack}`);
  }
}

function weeklySheetFindLargest(ss) {
  ss = ss || fetchSpreadsheet();
  const sheets = ss.getSheets();
  let weekNumbers = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();

    // Check if the sheet name starts with the specified prefix
    if (sheetName.startsWith(weeklySheetPrefix)) {
      // Extract the remaining text (the number) after the prefix
      const numberAsString = sheetName.replace(weeklySheetPrefix,"");

      // Convert the string to an integer and add it to our array
      const sheetNumber = parseInt(numberAsString);
      if (!isNaN(sheetNumber)) {
        weekNumbers.push(sheetNumber);
      }
    }
  }
  // If we found any sheets, return the largest number
  if (weekNumbers.length > 0) {
    let max = Math.max(...weekNumbers);
    let icon = max <= 10 ? numberMap[max] : (numberMap[Math.floor(max/10)] + numberMap[max - Math.floor(max/10)*10]);
    Logger.log(`✅ Found ${icon} as highest weekly sheet`)
    return max;
  } else {
    Logger.log(`⚠️ Unable to locate any weekly sheets at this point, returning 1`)
    return 1;
  }
}
/**
 * Robustly finds and extracts all relevant user and admin data from an existing sheet.
 * It dynamically finds the number of players and defensively checks for optional ranges.
 * @param {number} week The week number.
 * @param {object} forms The main forms data object.
 * @returns {object|null} An object with all preserved data, or null if essential ranges are missing.
 */
function getExistingWeeklySheetData(ss, week, forms) {
  const data = { playerData: {} };

  try {
    // Find the essential player data block. We use the NAMES range to find the top-left corner.
    const namesRangeStart = ss.getRangeByName(`NAMES_${week}`);
    const picksRange = ss.getRangeByName(`${LEAGUE}_PICKS_${week}`);
    
    if (!namesRangeStart || !picksRange) {
      Logger.log(`⚠️ Essential NAMES or PICKS named range not found. Cannot preserve data.`);
      return null;
    }

    // Now get all data using this dynamic row count.
    const names = ss.getRangeByName(`NAMES_${week}`).getValues().flat();
    Logger.log(`NAMES FOUND: ${names}`);
    const picks = ss.getRangeByName(`${LEAGUE}_PICKS_${week}`).getValues()

    // Defensively get optional player data.
    let tiebreakers = [];
    const tiebreakerRange = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}`);
    if (tiebreakerRange) {
      tiebreakers = tiebreakerRange.getValues();
      Logger.log(`⚖️ Tiebreaker range found; preserving tiebreakers.`);
    } else {
      Logger.log(`🚫 Tiebreaker range not found. Skipping preservation.`);
    }

    let tiebreakers2 = [];
    const tiebreaker2Range = ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}`);
    if (tiebreaker2Range) {
      tiebreakers2 = tiebreaker2Range.getValues();
      Logger.log(`⚖️ Tiebreaker 2 range found; preserving second tiebreakers.`);
    } else {
      Logger.log(`🚫 Tiebreaker 2 range not found. Skipping preservation.`);
    }

    let comments = [];
    const commentsRange = ss.getRangeByName(`COMMENTS_${week}`);
    if (commentsRange) {
      comments = commentsRange.getValues();
      Logger.log(`💬 Comments range found; preserving comments.`);
    } else {
      Logger.log(`🚫 Comments range not found. Skipping preservation.`);
    }

    // Use your mapping function to get the column order of the OLD games.
    data.oldMatchupMap = outcomeDataValidationMapping(week, forms, `${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    if (!data.oldMatchupMap) {
      Logger.log(`Could not create a game map from the old sheet. Aborting data preservation.`);
      return null;
    }

    // Map all scraped data to player names for easy re-sorting.
    names.forEach((name, index) => {
      if (name) {
        data.playerData[name] = {
          picks: picks[index] || [],
          tiebreaker: tiebreakers[index] ? tiebreakers[index][0] : '',
          tiebreaker2: tiebreakers2[index] ? tiebreakers2[index][0] : '',
          comment: comments[index] ? comments[index][0] : ''
        };
      }
    });

    Logger.log(`🧮 Gathering existing data for outcomes, margins, and spreads if present.`);
    // Get admin-entered data (also defensively).
    const outcomesRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const marginsRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`);
    const spreadsRange = ss.getRangeByName(`${LEAGUE}_SPREADS_${week}`);
    const tiebreakerOutcomeRange = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}_OUTCOME`);
    const tiebreaker2OutcomeRange = ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}_OUTCOME`);

    data.outcomes = outcomesRange ? outcomesRange.getValues()[0] : [];
    data.margins = marginsRange ? marginsRange.getValues()[0] : [];
    data.spreads = spreadsRange ? spreadsRange.getValues()[0] : [];
    data.tiebreaker = tiebreakerOutcomeRange ? tiebreakerOutcomeRange.getValue() : '';
    data.tiebreaker2 = tiebreaker2OutcomeRange ? tiebreaker2OutcomeRange.getValue() : '';

    Logger.log(`↩️ Returning collected data to the weekly sheet builder.`);
    return Object.keys(data.playerData).length > 0 ? data : null;

  } catch (err) {
    Logger.log(`⚠️ An error occurred during data scraping: ${err.stack}`);
    return null;
  }
}

/**
 * Remaps and repopulates preserved data onto a newly structured sheet.
 * Handles reordering of both players (rows) and games (columns).
 * @param {Sheet} sheet The target sheet object.
 * @param {number} week The week number.
 * @param {object} existingData The object returned by getExistingWeeklySheetData.
 * @param {object} newMatchupMap A map of { "AWAY @ HOME": columnIndex } for the new sheet layout.
 * @param {Array<string>} newMemberList The official new list of member names in the correct order.
 */
function remapAndRepopulateData(ss, week, existingData, newMatchupMap, newMemberList) {
  const { oldMatchupMap, playerData, outcomes, margins, spreads, tiebreaker, tiebreaker2 } = existingData;
  const unplacedMembers = [];
  
  // --- Part 1: Remap and Repopulate Player Data (Picks, Comments, etc.) ---
  const newPicks = new Array(newMemberList.length).fill(null).map(() => []);
  const newTiebreakers = new Array(newMemberList.length).fill(null).map(() => ['']);
  const newTiebreakers2 = new Array(newMemberList.length).fill(null).map(() => ['']);
  const newComments = new Array(newMemberList.length).fill(null).map(() => ['']);

  // Create a map for quick lookups of new player positions
  const newMemberMap = newMemberList.reduce((acc, name, index) => {
    acc[name] = index;
    return acc;
  }, {});
  
  // Find which old players still exist
  for (const playerName in playerData) {
    if (newMemberMap.hasOwnProperty(playerName)) {
      const newIndex = newMemberMap[playerName];
      const oldPlayer = playerData[playerName];
      
      // Place tiebreaker and comment directly
      newTiebreakers[newIndex][0] = oldPlayer.tiebreaker;
      newTiebreakers2[newIndex][0] = oldPlayer.tiebreaker2;
      newComments[newIndex][0] = oldPlayer.comment;
      
      // Now, remap the picks based on the new game order
      const reorderedPicks = [];
      for (const game in newMatchupMap) {
        const newColIndex = newMatchupMap[game]; // This is the 1-based column number
        
        if (oldMatchupMap.hasOwnProperty(game)) {
          const oldColIndex = oldMatchupMap[game];
          reorderedPicks[newColIndex - 1] = oldPlayer.picks[oldColIndex - 1] || '';
        } else {
          reorderedPicks[newColIndex - 1] = ''; // New game, so no old pick exists
        }
      }
      newPicks[newIndex] = reorderedPicks;
      
    } else {
      unplacedMembers.push(playerName);
    }
  }
  
  if (unplacedMembers.length > 0) {
    Logger.log(`⚠️ Unplaced members (not found in new list): ${unplacedMembers.join(', ')}`);
  }

  // Write the reordered player data to the sheet in batch
  ss.getRangeByName(`${LEAGUE}_PICKS_${week}`)?.setValues(newPicks);
  ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}`)?.setValues(newTiebreakers);
  ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}`)?.setValues(newTiebreakers2);
  ss.getRangeByName(`COMMENTS_${week}`)?.setValues(newComments);

  // --- Part 2: Remap and Repopulate Admin Data (Outcomes, Spreads, etc.) ---
  const newOutcomes = [];
  const newMargins = [];
  const newSpreads = [];

  for (const game in newMatchupMap) {
    const newColIndex = newMatchupMap[game];
    if (oldMatchupMap.hasOwnProperty(game)) {
      const oldColIndex = oldMatchupMap[game];
      newOutcomes[newColIndex - 1] = outcomes[oldColIndex - 1] || '';
      newMargins[newColIndex - 1] = margins[oldColIndex - 1] || '';
      newSpreads[newColIndex - 1] = spreads[oldColIndex - 1] || '';
    } else {
      // Initialize empty values for new games
      newOutcomes[newColIndex - 1] = '';
      newMargins[newColIndex - 1] = '';
      newSpreads[newColIndex - 1] = '';
    }
  }
  
  // Write the reordered admin data in batch
  ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`)?.setValues([newOutcomes]);
  ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`)?.setValues([newMargins]);
  ss.getRangeByName(`${LEAGUE}_SPREADS_${week}`)?.setValues([newSpreads]);
  ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}_OUTCOME`)?.setValue(tiebreaker);
  ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}_OUTCOME`)?.setValue(tiebreaker2);
}


// RANDOM - random integer function for selecting Game of the Week
function getRandomInt(min, max) {
      min = Math.ceil(min);
      max = Math.floor(max);
      return Math.floor(Math.random() * (max - min + 1)) + min;
}

// WEEKLY SHEET SUPPORT IN-CELL FORMULAS

/**
 * Calculates the "True Win Probability" for a pick'em pool, where ties dilute the chance of winning.
 * The sum of all players' chances will now equal 100%.
 *
 * @param {range} playerPicksRange The range of all players' picks (e.g., E3:T17).
 * @param {range} resultsRange The range of the actual game outcomes (e.g., E19:T19).
 * @param {range} currentScoresRange The range of the players' current scores (e.g., B3:B17).
 * @param {range} [bonusRange] Optional. A row of bonus multipliers for each game (e.g., E21:T21).
 * @param {range} [winnersRange] Optional. A column that marks the winner(s) of the week.
 * @param {range} [spreadInfoRange] Optional. A row where each cell contains the favorite and spread (e.g., "HOU -2.5").
 * @return {Array<Array<number | string>>} A column of true win probabilities.
 * @customfunction
 */
function calculateWinProbability(playerPicksRange, resultsRange, currentScoresRange, bonusRange, winnersRange, spreadInfoRange) {
  
  if (winnersRange && winnersRange.flat().some(cell => cell)) {
    // This logic remains the same, as it's a manual override.
    const winnersCount = winnersRange.flat().filter(cell => cell).length;
    const winShare = winnersCount > 0 ? 1.0 / winnersCount : 0;
    return winnersRange.map(row => row[0] ? [winShare] : [0.0]);
  }
  
  const originalNumPlayers = playerPicksRange.length;
  const activePlayers = [];
  
  for (let i = 0; i < originalNumPlayers; i++) {
    if (playerPicksRange[i].some(pick => pick)) {
      activePlayers.push({
        originalIndex: i,
        picks: playerPicksRange[i],
        score: currentScoresRange[i][0]
      });
    }
  }

  if (activePlayers.length === 0) {
    return new Array(originalNumPlayers).fill([""]);
  }

  const activePlayerPicks = activePlayers.map(p => p.picks);
  const activeCurrentScores = activePlayers.map(p => p.score);
  const numActivePlayers = activePlayers.length;

  const allPlayersTied = activeCurrentScores.every(score => score === activeCurrentScores[0]);
  if (allPlayersTied) {
    const fullOutput = new Array(originalNumPlayers).fill([""]);
    // UPDATED LOGIC #1: Split the win chance evenly if all are tied
    const winShare = 1.0 / numActivePlayers;
    activePlayers.forEach(player => {
      fullOutput[player.originalIndex] = [winShare];
    });
    return fullOutput;
  }

  const gameResults = resultsRange[0];
  const numGames = gameResults.length;
  const useSpread = Array.isArray(spreadInfoRange);
  const bonuses = Array.isArray(bonusRange) ? bonusRange[0] : new Array(numGames).fill(1);
  const spreadInfo = useSpread ? spreadInfoRange[0] : [];
  const relevantGames = [];
  
  // Helper Function for Normal Distribution CDF (approximates NORMSDIST)
  function standardNormalCdf(z) {
    const p = 0.3275911, a1 = 0.254829592, a2 = -0.284496736, a3 = 1.421413741, a4 = -1.453152027, a5 = 1.061405429;
    const sign = (z >= 0) ? 1 : -1;
    const t = 1.0 / (1.0 + p * Math.abs(z));
    const poly = a1*t + a2*Math.pow(t,2) + a3*Math.pow(t,3) + a4*Math.pow(t,4) + a5*Math.pow(t,5);
    const erf = 1 - poly * Math.exp(-Math.pow(z, 2));
    return 0.5 * (1.0 + sign * erf);
  }

  for (let j = 0; j < numGames; j++) {
    if (gameResults[j] === "") {
      const picksForGame = new Set(activePlayerPicks.map(row => row[j]).filter(pick => pick));
      if (picksForGame.size === 2) {
        const outcomes = Array.from(picksForGame);
        const gameInfo = { columnIndex: j, outcomes: outcomes, bonus: bonuses[j] || 1 };
        let spreadDataParsed = false;
        
        if (useSpread && spreadInfo[j] && typeof spreadInfo[j] === 'string') {
          const cleanedSpreadText = spreadInfo[j].replace(/–|—/g, '-').trim();
          const match = cleanedSpreadText.match(/^([A-Z]{2,3})\s+(-?\d+(\.\d+)?)$/);
          
          if (match) {
            const favoriteTeam = match[1];
            const spreadValue = parseFloat(match[2]);
            const underdogTeam = outcomes.find(team => team !== favoriteTeam);
            
            if (underdogTeam && spreadValue <= 0) {
              const favoriteProb = standardNormalCdf(-spreadValue / 13.86); // This line can now execute correctly
              gameInfo.probabilities = { [favoriteTeam]: favoriteProb, [underdogTeam]: 1 - favoriteProb };
              spreadDataParsed = true;
            }
          }
        }
        
        if (!spreadDataParsed) {
          gameInfo.probabilities = { [outcomes[0]]: 0.5, [outcomes[1]]: 0.5 };
        }
        relevantGames.push(gameInfo);
      }
    }
  }
  let finalProbabilities;
  const numRelevantGames = relevantGames.length;

  if (numRelevantGames === 0) {
    const maxScore = Math.max(...activeCurrentScores);
    // UPDATED LOGIC #2: If week is over, split win chance among tied leaders
    const winnersCount = activeCurrentScores.filter(score => score === maxScore).length;
    const winShare = winnersCount > 0 ? 1.0 / winnersCount : 0;
    finalProbabilities = activeCurrentScores.map(score => (score === maxScore ? winShare : 0.0));
  } else if (numRelevantGames > 18) {
     finalProbabilities = new Array(numActivePlayers).fill("Calculation too complex (>18 games)");
  } else {
    const totalPermutations = Math.pow(2, numRelevantGames);
    const winProbabilitySum = new Array(numActivePlayers).fill(0.0);

    for (let i = 0; i < totalPermutations; i++) {
      let tempFinalScores = [...activeCurrentScores];
      let permutationProbability = 1.0;
      let binaryPermutation = i.toString(2).padStart(numRelevantGames, '0');
      for (let j = 0; j < numRelevantGames; j++) {
        const game = relevantGames[j];
        const winningTeam = game.outcomes[parseInt(binaryPermutation[j])];
        permutationProbability *= game.probabilities[winningTeam];
        for (let p = 0; p < numActivePlayers; p++) {
          if (activePlayerPicks[p][game.columnIndex] === winningTeam) {
            tempFinalScores[p] += game.bonus;
          }
        }
      }
      
      const maxScoreInPermutation = Math.max(...tempFinalScores);
      
      // --- UPDATED LOGIC #3: Find all winners and distribute the probability ---
      const winnersInPermutation = [];
      for (let p = 0; p < numActivePlayers; p++) {
        if (tempFinalScores[p] === maxScoreInPermutation) {
          winnersInPermutation.push(p);
        }
      }
      
      const numWinners = winnersInPermutation.length;
      if (numWinners > 0) {
        const distributedProbability = permutationProbability / numWinners;
        for (const winnerIndex of winnersInPermutation) {
          winProbabilitySum[winnerIndex] += distributedProbability;
        }
      }
    }
    finalProbabilities = winProbabilitySum;
  }

  const fullOutput = new Array(originalNumPlayers).fill([""]);
  activePlayers.forEach((player, i) => {
    fullOutput[player.originalIndex] = [finalProbabilities[i]];
  });
  return fullOutput;
}


/**
 * Calculates a "Wildcard Score" for each player, measuring how much their picks deviate from the group consensus.
 * A score of 0% means the player picked with the majority on every game.
 * A score of 100% means the player picked with the minority on every game.
 * Unanimous picks are now correctly included to establish a true 0-to-100 scale.
 *
 * @param {range} playerPicksRange The range containing all player picks (e.g., E3:T17).
 * @return {Array<Array<number | string>>} A column of wildcard scores from 0.0 to 1.0 (format as % in Sheets).
 * @customfunction
 */
function calculateWildcardScore(playerPicksRange) {
  if (!playerPicksRange || playerPicksRange.length === 0) {
    return [[""]];
  }

  const numPlayers = playerPicksRange.length;
  const numGames = playerPicksRange[0].length;

  // --- Step 1: Analyze the Consensus for Each Game ---
  const consensusData = [];
  for (let j = 0; j < numGames; j++) {
    const pickCounts = {};
    let totalPicksInGame = 0;

    for (let i = 0; i < numPlayers; i++) {
      const pick = playerPicksRange[i][j];
      if (pick) { // Only count non-empty picks
        pickCounts[pick] = (pickCounts[pick] || 0) + 1;
        totalPicksInGame++;
      }
    }
    
    let majorityPopularity = 0.5;
    let minorityPopularity = 0.5;
    const pickPopularity = {};

    if (totalPicksInGame > 0) {
      const uniquePicks = Object.keys(pickCounts);

      // Calculate popularity for each picked team
      uniquePicks.forEach(team => {
        pickPopularity[team] = pickCounts[team] / totalPicksInGame;
      });

      if (uniquePicks.length === 1) {
        // Unanimous pick - this is the corrected logic
        majorityPopularity = 1.0;
        minorityPopularity = 0.0;
      } else if (uniquePicks.length > 1) {
        // Contested pick
        const popularities = Object.values(pickPopularity);
        majorityPopularity = Math.max(...popularities);
        minorityPopularity = Math.min(...popularities);
      }
    }

    consensusData.push({
      popularity: pickPopularity,
      // Boldness for picking the majority (safest pick)
      minBoldness: 1 - majorityPopularity,
      // Boldness for picking the minority (riskiest pick)
      maxBoldness: 1 - minorityPopularity
    });
  }

  // --- Step 2: Calculate Scores for Each Player ---
  const finalScores = [];
  for (let i = 0; i < numPlayers; i++) {
    const playerRow = playerPicksRange[i];
    let rawWildcardScore = 0;
    let minPossibleScore = 0; // The score for picking all favorites
    let maxPossibleScore = 0; // The score for picking all underdogs

    if (playerRow.every(pick => !pick)) {
      finalScores.push([""]); // Return empty for empty rows
      continue;
    }

    for (let j = 0; j < numGames; j++) {
      const pick = playerRow[j];
      const gameConsensus = consensusData[j];

      if (pick) { // Only score games where a pick was made
        const popularityOfPick = gameConsensus.popularity[pick] || 0;
        rawWildcardScore += (1 - popularityOfPick);
        minPossibleScore += gameConsensus.minBoldness;
        maxPossibleScore += gameConsensus.maxBoldness;
      }
    }

    // --- Step 3: Normalize the Score ---
    const scoreRange = maxPossibleScore - minPossibleScore;
    
    if (scoreRange === 0) {
      // This case now only happens if all games are perfect 50/50 splits.
      // In this scenario, no "bold" path exists, so a 0% score is appropriate.
      finalScores.push([0]);
      continue;
    }

    const normalizedScore = (rawWildcardScore - minPossibleScore) / scoreRange;
    finalScores.push([normalizedScore]);
  }

  return finalScores;
}


// ============================================================================================================================================
// WEEKLY OPERATIONS
// ============================================================================================================================================

/**
 * Maps a sheet week number onto the ESPN scoreboard's week/season-type pair.
 * Regular season weeks pass straight through; the playoffs restart their numbering.
 */
function scoreboardWeekParams(week) {
  const w = Number(week);
  if (w <= REGULAR_SEASON) return { week: w, seasontype: 2 };
  const playoffRound = { 19: 1, 20: 2, 21: 3, 23: 5 }; // 4 is the Pro Bowl, which we never score
  return { week: playoffRound[w] || 1, seasontype: 3 };
}

/**
 * Pulls outcomes for one week straight from the API, with no sidebar.
 *
 * @param {number} week Week to score.
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {Object} {placed, completed, total} counts.
 */
function importOutcomesForWeek(week, ss) {
  ss = fetchSpreadsheet(ss);
  const formsData = fetchProperties('forms') || {};
  if (!formsData[week] || !formsData[week].gamePlan) {
    return { placed: false, completed: 0, total: 0, reason: `no form was ever built for week ${week}` };
  }
  const params = scoreboardWeekParams(week);
  const payload = espnFetchJson(`${SCOREBOARD}?week=${params.week}&seasontype=${params.seasontype}`);
  const events = (payload && payload.events) || [];
  const analysis = parseApiEvents(events, formsData[week].gamePlan);
  const total = formsData[week].gamePlan.games.length;
  if (!analysis.complete.length) {
    return { placed: false, completed: 0, total: total, reason: `no games have finished yet` };
  }
  updateSheetsWithApiOutcomes(ss, week, analysis.complete, formsData, true);
  return { placed: true, completed: analysis.complete.length, total: total };
}

/**
 * The whole end-of-week routine in one action: pull outcomes, refresh formulas, then
 * recompute every payout tab. Refuses to pay out a week that is not fully scored.
 */
function closeOutWeek() {
  const ss = fetchSpreadsheet();
  const ui = fetchUi();
  const formsData = fetchProperties('forms') || {};
  const built = Object.keys(formsData).map(Number).filter(w => !isNaN(w)).sort((a, b) => a - b);
  if (!built.length) {
    ui.alert(`⛔ NO WEEKS YET`, `No weekly forms have been built, so there is nothing to close out.`, ui.ButtonSet.OK);
    return;
  }

  const suggested = built[built.length - 1];
  const prompt = ui.prompt(`✅ CLOSE OUT WEEK`,
    `Which week should I close out?\n\nThis pulls the final scores, refreshes the formulas, and recomputes every payout tab.\n\nWeeks with forms: ${built.join(', ')}\n\nEnter a week number (or leave blank for ${suggested}):`,
    ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  const entered = prompt.getResponseText().trim();
  const week = entered ? parseInt(entered, 10) : suggested;
  if (isNaN(week) || built.indexOf(week) === -1) {
    ui.alert(`⚠️ UNKNOWN WEEK`, `Week ${entered} has no form. Weeks with forms: ${built.join(', ')}`, ui.ButtonSet.OK);
    return;
  }

  const notes = [];
  try {
    ss.toast(`Fetching final scores for week ${week}...`, `1/3 OUTCOMES`);
    const outcome = importOutcomesForWeek(week, ss);
    notes.push(outcome.placed
      ? `✅ Outcomes: ${outcome.completed} of ${outcome.total} games final`
      : `⏩ Outcomes: ${outcome.reason}`);

    ss.toast(`Refreshing formulas...`, `2/3 FORMULAS`);
    allFormulasUpdate(ss);
    notes.push(`✅ Formulas refreshed`);

    ss.toast(`Recomputing payouts...`, `3/3 PAYOUTS`);
    const data = computePayouts(ss);
    writeWeeklyPlaces(ss, data);
    wklyPayoutSheet(ss, data);
    wklyConsensusPayoutSheet(ss, data);
    seasonPointsSheet(ss, data);
    seasonTotalsSheet(ss, data);
    if (data.playoffWeeksScored) playoffPointsSheet(ss, data);
    summaryPayoutSheet(ss, data);
    standingsSheet(ss, data);
    wklyRankSheet(ss, data);
    seasonRankSheet(ss, data);
    rankJumpChartSheet(ss, data);

    // Only a fully scored week pays out, so say plainly whether this one did
    if (data.weeklyMoney[week]) {
      const winners = Object.keys(data.weeklyPlaces[week]).filter(n => data.weeklyPlaces[week][n] === 1);
      const paid = Object.keys(data.weeklyMoney[week]).filter(n => data.weeklyMoney[week][n] > 0).length;
      notes.push(`✅ Week ${week} paid out: ${winners.join(', ')} took first, ${paid} members in the money`);
      const consensus = data.consensusMoney[week];
      if (consensus) {
        notes.push(consensus.refunded
          ? `↩️ Consensus: nobody beat the crowd, pot refunded evenly`
          : `✅ Consensus: ${consensus.winners} beat the crowd`);
      }
    } else {
      notes.push(`⏸️ Week ${week} is not fully scored yet, so no money was paid. Enter any missing outcomes on the WK${week} sheet and run this again.`);
    }

    ui.alert(`✅ WEEK ${week} CLOSE-OUT`, notes.join(`\n\n`), ui.ButtonSet.OK);
  } catch (err) {
    Logger.log(`❌ closeOutWeek failed: ${err.stack}`);
    ui.alert(`⚠️ CLOSE-OUT FAILED`, `${notes.join('\n')}\n\nThen it stopped:\n${err.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Checks the things that quietly drift over a long season and reports anything wrong.
 */
function healthCheck() {
  const ss = fetchSpreadsheet();
  const ui = fetchUi();
  const config = fetchProperties('configuration') || {};
  const memberData = fetchProperties('members') || { memberOrder: [], members: {} };
  const formsData = fetchProperties('forms') || {};
  const problems = [], notes = [];

  // --- roster
  const roster = memberData.memberOrder.map(id => memberData.members[id] && memberData.members[id].name).filter(Boolean);
  notes.push(`👥 ${roster.length} members on the roster`);
  const dupes = roster.filter((n, i) => roster.map(x => x.toLowerCase()).indexOf(n.toLowerCase()) !== i);
  if (dupes.length) problems.push(`Duplicate member names: ${dupes.join(', ')}`);
  const noEmail = memberData.memberOrder.filter(id => !(memberData.members[id] || {}).email).length;
  if (noEmail) notes.push(`📧 ${noEmail} member(s) have no email on file`);

  // --- configuration matches the pool rules
  if (config.pickemsAts) problems.push(`Against-the-spread is ON but this pool picks straight up`);
  if (!config.bonusInclude) problems.push(`Bonus is OFF, so the weekly bonus game cannot be set`);
  if (!config.tiebreakerInclude) problems.push(`Tiebreakers are OFF but the weekly pool needs them to order places`);
  if (config.survivorInclude || config.eliminatorInclude) problems.push(`Survivor/Eliminator is ON but this pool does not run them`);
  if (!config.kickoffLock) notes.push(`🔓 Kickoff lock is off, so late picks are still accepted`);

  // --- per-week named ranges
  const weeks = Object.keys(formsData).map(Number).filter(w => !isNaN(w)).sort((a, b) => a - b);
  const required = ['NAMES_%', `${LEAGUE}_PICKS_%`, `${LEAGUE}_PICKEM_OUTCOMES_%`, `${LEAGUE}_BONUS_%`,
                    `${LEAGUE}_TIEBREAKER_%`, `${LEAGUE}_TIEBREAKER2_%`, `${LEAGUE}_CONSENSUS_%`, 'WIN_%'];
  weeks.forEach(week => {
    if (!ss.getSheetByName(`${weeklySheetPrefix}${week}`)) return; // sheet not built yet is fine
    const missing = required.filter(r => !ss.getRangeByName(r.replace('%', week)));
    if (missing.length) problems.push(`WK${week} is missing named ranges: ${missing.map(m => m.replace('%', week)).join(', ')}`);
  });

  // --- money reconciles
  try {
    const data = computePayouts(ss);
    const scored = data.weeksScored.length;
    notes.push(`📅 ${scored} week(s) fully scored`);
    notes.push(`💵 Weekly pot $${data.weeklyPot.toFixed(2)}, consensus pot $${data.consensusPot.toFixed(2)}, ${data.ladder.length} places at ${data.ladder.join('/')}%`);
    const paidWeekly = Object.keys(data.totals.weeklyWins).reduce((s, n) => s + data.totals.weeklyWins[n], 0);
    const owedWeekly = data.weeklyPot * scored;
    if (Math.abs(paidWeekly - owedWeekly) > 0.005) {
      problems.push(`Weekly Wins paid $${paidWeekly.toFixed(2)} but collected $${owedWeekly.toFixed(2)} across ${scored} scored week(s)`);
    }
    const paidConsensus = Object.keys(data.totals.weeklyConsensus).reduce((s, n) => s + data.totals.weeklyConsensus[n], 0);
    const owedConsensus = data.consensusPot * Object.keys(data.consensusMoney).length;
    if (Math.abs(paidConsensus - owedConsensus) > 0.005) {
      problems.push(`Weekly Consensus paid $${paidConsensus.toFixed(2)} but collected $${owedConsensus.toFixed(2)}`);
    }
    // every member should appear on every payout tab
    ['Wkly Payout', 'Wkly Consensus', 'Season Points', 'Summary Payout'].forEach(tab => {
      const sheet = ss.getSheetByName(tab);
      if (!sheet) { notes.push(`📄 ${tab} not built yet`); return; }
      const listed = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat()
        .map(v => v ? v.toString() : '');
      const absent = roster.filter(n => listed.indexOf(n) === -1);
      if (absent.length) problems.push(`${tab} is missing: ${absent.join(', ')} -- re-run Update Payouts`);
    });
  } catch (err) {
    problems.push(`Could not reconcile payouts: ${err.message}`);
  }

  const verdict = problems.length
    ? `⚠️ ${problems.length} PROBLEM${problems.length === 1 ? '' : 'S'} FOUND`
    : `✅ ALL CHECKS PASSED`;
  ui.alert(verdict, (problems.length ? problems.map(p => `❌ ${p}`).join(`\n`) + `\n\n---\n\n` : ``) + notes.join(`\n`), ui.ButtonSet.OK);
  Logger.log(`🩺 Health check: ${problems.length} problem(s). ${problems.join(' | ')}`);
}

/**
 * Duplicates the whole spreadsheet into the pool's Drive folder, timestamped.
 * Worth running before anything that rebuilds sheets.
 */
function snapshotSpreadsheet(silent) {
  const ss = fetchSpreadsheet();
  const config = fetchProperties('configuration') || {};
  try {
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd HH-mm');
    const name = `${ss.getName()} -- BACKUP ${stamp}`;
    const copy = DriveApp.getFileById(ss.getId()).makeCopy(name);
    try {
      copy.moveTo(getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`));
    } catch (err) {
      Logger.log(`⚠️ Snapshot created but could not be filed in the pool folder: ${err.message}`);
    }
    Logger.log(`📸 Snapshot created: ${name}`);
    if (!silent) {
      showLinkDialog(copy.getUrl(), `📸 Snapshot Created`, name,
        `\nA full copy of this spreadsheet has been saved to your pool folder. Nothing in this file changed.`);
    }
    return copy;
  } catch (err) {
    Logger.log(`❌ Snapshot failed: ${err.stack}`);
    if (!silent) fetchUi().alert(`⚠️ SNAPSHOT FAILED`, `Could not copy the spreadsheet:\n\n${err.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Turns the kickoff lock on or off. When on, each new form gets a one-time trigger that
 * closes it at the week's first kickoff so nobody submits picks after games start.
 */
function toggleKickoffLock() {
  const ui = fetchUi();
  const config = fetchProperties('configuration') || {};
  const turningOn = !config.kickoffLock;
  const answer = ui.alert(turningOn ? `🔒 ENABLE KICKOFF LOCK` : `🔓 DISABLE KICKOFF LOCK`,
    turningOn
      ? `Each new weekly form will close automatically at that week's first kickoff, so late picks are refused.\n\nForms already created are unaffected.\n\nEnable it?`
      : `New forms will stay open after kickoff, accepting late picks.\n\nDisable the lock?`,
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  config.kickoffLock = turningOn;
  saveProperties('configuration', config);
  fetchSpreadsheet().toast(turningOn ? `Kickoff lock enabled` : `Kickoff lock disabled`, turningOn ? `🔒 LOCKED` : `🔓 UNLOCKED`);
  Logger.log(`${turningOn ? '🔒' : '🔓'} kickoffLock set to ${turningOn}`);
}

/**
 * Earliest kickoff in a week, built from the game date plus its hour and minute.
 *
 * Ben's form-lock trigger uses new Date(game.date), which is midnight on the game's day
 * rather than the actual kickoff. Reminders need the real time, so hour/minute are applied here.
 *
 * @param {Object} gamePlan The week's saved game plan.
 * @returns {Date|null} The first kickoff, or null if it cannot be determined.
 */
function weekFirstKickoff(gamePlan) {
  if (!gamePlan || !gamePlan.games || !gamePlan.games.length) return null;
  let earliest = null;
  gamePlan.games.forEach(game => {
    const base = new Date(game.date);
    if (isNaN(base.getTime())) return;
    const when = new Date(base.getFullYear(), base.getMonth(), base.getDate(),
      Number(game.hour) || 0, Number(game.minute) || 0);
    if (!earliest || when < earliest) earliest = when;
  });
  return earliest;
}

/**
 * Works out who still owes picks for a week.
 *
 * The stored respondent list is refreshed from the form first -- otherwise anyone who
 * submitted since the last import would be nagged for picks they have already made.
 *
 * @param {number} week Week to check.
 * @returns {Object} {form, targets, missingEmail, pendingCount} or {error}
 */
function reminderRecipients(week) {
  try {
    syncFormResponses(week);        // make sure the outstanding list is current
  } catch (err) {
    Logger.log(`⚠️ Could not refresh responses for week ${week} before reminding: ${err.message}`);
  }

  const formsData = fetchProperties('forms') || {};
  const memberData = fetchProperties('members') || { memberOrder: [], members: {} };
  const form = formsData[week];
  if (!form) return { error: `There is no form recorded for week ${week}.` };

  const pending = form.nonRespondents || [];
  const targets = [], missingEmail = [];
  pending.forEach(id => {
    const m = memberData.members[id];
    if (!m) return;
    if (m.email) targets.push({ name: m.name, email: m.email });
    else missingEmail.push(m.name);
  });
  return { form: form, targets: targets, missingEmail: missingEmail, pendingCount: pending.length };
}

/**
 * Sends the reminder emails and records that the week has been reminded.
 *
 * @param {number} week Week being reminded about.
 * @param {Object} info Result of reminderRecipients().
 * @returns {Object} {sent, failures}
 */
function sendReminderEmails(week, info) {
  const config = fetchProperties('configuration') || {};
  const url = info.form.publishedUrl || '';
  const subject = `${config.groupName || `${LEAGUE} Picks`} -- Week ${week} picks needed`;

  let sent = 0;
  const failures = [];
  info.targets.forEach(t => {
    try {
      MailApp.sendEmail(t.email, subject,
        `${t.name},\n\nYour week ${week} picks are not in yet.\n\n${url}\n\nGood luck.\n`);
      sent++;
    } catch (err) {
      failures.push(`${t.name} (${err.message})`);
    }
  });

  // Recorded so a re-fired trigger cannot email the same week twice
  const formsData = fetchProperties('forms') || {};
  if (formsData[week]) {
    formsData[week].reminderSentAt = new Date().toISOString();
    formsData[week].reminderSentCount = sent;
    saveProperties('forms', formsData);
  }

  Logger.log(`📧 Week ${week}: sent ${sent} reminder(s). Failures: ${failures.join('; ') || 'none'}`);
  return { sent: sent, failures: failures };
}

/**
 * Menu action: remind whoever has not submitted, after showing you the list.
 */
function remindNonRespondents() {
  const ss = fetchSpreadsheet();
  const ui = fetchUi();
  const formsData = fetchProperties('forms') || {};

  const openWeeks = Object.keys(formsData).map(Number)
    .filter(w => !isNaN(w) && formsData[w].active !== false).sort((a, b) => b - a);
  const week = openWeeks.length ? openWeeks[0] : null;
  if (!week) {
    ui.alert(`⛔ NO OPEN FORM`, `There is no active weekly form to remind anyone about.`, ui.ButtonSet.OK);
    return;
  }

  ss.toast(`Checking who still owes picks for week ${week}...`, `📧 CHECKING`);
  const info = reminderRecipients(week);
  if (info.error) {
    ui.alert(`⚠️ CANNOT REMIND`, info.error, ui.ButtonSet.OK);
    return;
  }
  if (!info.pendingCount) {
    ui.alert(`✅ EVERYONE IS IN`, `All members have submitted picks for week ${week}.`, ui.ButtonSet.OK);
    return;
  }
  if (!info.targets.length) {
    ui.alert(`⚠️ NO EMAIL ADDRESSES`, `${info.pendingCount} member(s) have not submitted, but none of them have an email on file:\n\n${info.missingEmail.join(', ')}\n\nAdd addresses via the sign-up form or the Member Manager.`, ui.ButtonSet.OK);
    return;
  }

  const preview = info.targets.slice(0, 20).map(t => `  • ${t.name} <${t.email}>`).join(`\n`)
    + (info.targets.length > 20 ? `\n  ...and ${info.targets.length - 20} more` : ``);
  const already = info.form.reminderSentAt
    ? `\nNote: a reminder for this week already went out on ${new Date(info.form.reminderSentAt).toLocaleString()}.\n` : ``;

  const answer = ui.alert(`📧 SEND ${info.targets.length} REMINDER${info.targets.length === 1 ? '' : 'S'}?`,
    `Week ${week} picks are still missing from:\n\n${preview}\n\n` +
    (info.missingEmail.length ? `No email on file for: ${info.missingEmail.join(', ')}\n\n` : ``) +
    already +
    `Each will get the form link. Mail is sent from your own Google account.\n\nSend now?`,
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) {
    ss.toast(`No reminders sent`, `⛔ CANCELED`);
    return;
  }

  const result = sendReminderEmails(week, info);
  ui.alert(`📧 ${result.sent} REMINDER${result.sent === 1 ? '' : 'S'} SENT`,
    `Week ${week} reminders sent to ${result.sent} member(s).` +
    (result.failures.length ? `\n\nFailed: ${result.failures.join(', ')}` : ``) +
    (info.missingEmail.length ? `\n\nSkipped, no email on file: ${info.missingEmail.join(', ')}` : ``),
    ui.ButtonSet.OK);
}

/**
 * Schedules the automatic reminder for a week: REMINDER_HOURS_BEFORE_KICKOFF hours before
 * that week's first kickoff. Called when a form is built, and safe to call again -- any
 * previously scheduled reminder for the same week is replaced.
 *
 * @param {number} week The week the form covers.
 * @param {Object} gamePlan That week's game plan.
 */
function scheduleWeeklyReminder(week, gamePlan) {
  const config = fetchProperties('configuration') || {};
  if (config.autoRemind === false) {
    Logger.log(`🔕 Auto-reminders are switched off; nothing scheduled for week ${week}.`);
    return;
  }

  const kickoff = weekFirstKickoff(gamePlan);
  if (!kickoff) {
    Logger.log(`⚠️ Could not determine the first kickoff for week ${week}; no reminder scheduled.`);
    return;
  }
  const sendAt = new Date(kickoff.getTime() - (REMINDER_HOURS_BEFORE_KICKOFF * 60 * 60 * 1000));
  if (sendAt <= new Date()) {
    Logger.log(`⏩ Week ${week} reminder time (${sendAt.toLocaleString()}) has already passed; nothing scheduled.`);
    return;
  }

  const formsData = fetchProperties('forms') || {};
  // Replace rather than stack, so rebuilding a form does not double-email
  if (formsData[week] && formsData[week].reminderTriggerId) {
    deleteTriggerById(formsData[week].reminderTriggerId);
  }

  const trigger = ScriptApp.newTrigger('executeWeeklyReminder').timeBased().at(sendAt).create();
  const triggerId = trigger.getUniqueId();
  PropertiesService.getDocumentProperties()
    .setProperty('triggerMeta_' + triggerId, JSON.stringify({ week: week, kind: 'reminder' }));

  if (formsData[week]) {
    formsData[week].reminderTriggerId = triggerId;
    formsData[week].reminderScheduledFor = sendAt.toISOString();
    delete formsData[week].reminderSentAt;      // a fresh form gets a fresh reminder
    saveProperties('forms', formsData);
  }

  Logger.log(`📧 Week ${week} reminder scheduled for ${sendAt.toLocaleString()} (${REMINDER_HOURS_BEFORE_KICKOFF}h before ${kickoff.toLocaleString()}).`);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Non-respondents will be emailed ${REMINDER_HOURS_BEFORE_KICKOFF} hours before kickoff`, `📧 AUTO-REMINDER SET`);
}

/**
 * Trigger handler for the automatic reminder. Sends nothing if the form has closed, if
 * everyone is in, or if this week has already been reminded.
 */
function executeWeeklyReminder(e) {
  const triggerId = e && e.triggerUid;
  const docProps = PropertiesService.getDocumentProperties();
  const meta = triggerId ? docProps.getProperty('triggerMeta_' + triggerId) : null;
  if (!meta) {
    Logger.log(`⭕ executeWeeklyReminder fired without usable metadata; cleaning up.`);
    if (triggerId) deleteTriggerById(triggerId);
    return;
  }

  const week = JSON.parse(meta).week;
  try {
    const formsData = fetchProperties('forms') || {};
    const form = formsData[week];
    if (!form) { Logger.log(`⭕ Week ${week} has no form; no reminder sent.`); return; }
    if (form.active === false) { Logger.log(`⏩ Week ${week} form is closed; no reminder sent.`); return; }
    if (form.reminderSentAt) { Logger.log(`⏩ Week ${week} was already reminded at ${form.reminderSentAt}; not sending again.`); return; }

    const info = reminderRecipients(week);
    if (info.error) { Logger.log(`⚠️ ${info.error}`); return; }
    if (!info.targets.length) {
      Logger.log(`✅ Week ${week}: nobody to remind (${info.pendingCount} outstanding, ${info.missingEmail.length} without an email).`);
      return;
    }
    const result = sendReminderEmails(week, info);
    Logger.log(`📧 Automatic week ${week} reminder: ${result.sent} sent, ${info.missingEmail.length} skipped for no email.`);
  } catch (err) {
    Logger.log(`❌ Automatic week ${week} reminder failed: ${err.stack}`);
  } finally {
    deleteTriggerById(triggerId);
    docProps.deleteProperty('triggerMeta_' + triggerId);
  }
}

/**
 * Turns automatic reminders on or off for future forms.
 */
function toggleAutoRemind() {
  const ui = fetchUi();
  const config = fetchProperties('configuration') || {};
  // Undefined means on, so that a pool set up before this feature existed still gets reminders
  const currentlyOn = config.autoRemind !== false;
  const turningOn = !currentlyOn;
  const answer = ui.alert(turningOn ? `📧 ENABLE AUTO-REMINDERS` : `🔕 DISABLE AUTO-REMINDERS`,
    turningOn
      ? `Each new weekly form will email anyone who has not submitted, ${REMINDER_HOURS_BEFORE_KICKOFF} hours before that week's first kickoff.\n\nThe respondent list is refreshed first, so nobody who has already picked gets a reminder.\n\nEnable it?`
      : `New forms will not send automatic reminders. You can still send them by hand from the menu.\n\nAlready-scheduled reminders for existing forms will still fire.\n\nDisable?`,
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  config.autoRemind = turningOn;
  saveProperties('configuration', config);
  fetchSpreadsheet().toast(turningOn ? `Automatic reminders enabled` : `Automatic reminders disabled`,
    turningOn ? `📧 AUTO-REMIND ON` : `🔕 AUTO-REMIND OFF`);
  Logger.log(`${turningOn ? '📧' : '🔕'} autoRemind set to ${turningOn}`);
}

// ============================================================================================================================================
// PAYOUTS
// ============================================================================================================================================
//
// One engine drives every pool's money so the tie rule lives in exactly one place.
//
// Two DIFFERENT tie rules are in play, and they must not be confused:
//   * MONEY (this section): ties consume the places below them and split the combined
//     percentage. A 3-way tie for 1st takes the percentages for places 1, 2 and 3, adds
//     them together and splits the total three ways -- the next finisher is then 4th.
//   * SEASON POINTS (weeklyPlacementPoints): tied players all receive that tier's full
//     points, and the next DISTINCT score receives the next tier down. Nothing is consumed.

/**
 * How many places get paid, given the size of the pool.
 *
 * @param {number} entries Number of paid entries.
 * @returns {Array<number>} Percentage of the pot for each place, best first.
 */
function payoutLadder(entries) {
  const count = Number(entries) || 0;
  const rule = PAYOUT_PLACE_BREAKPOINTS.find(b => count >= b.minEntries) || { places: 3 };
  return PAYOUT_LADDERS[rule.places] || PAYOUT_LADDERS[3];
}

/**
 * Standard competition ranking: ties share the better rank and the next distinct value skips
 * ahead (1, 2, 2, 4). Higher values rank better.
 *
 * @param {Object} valuesByName Map of name to the value being ranked.
 * @returns {Object} Map of name to rank.
 */
function competitionRanks(valuesByName) {
  const names = Object.keys(valuesByName || {});
  const sorted = names.slice().sort((a, b) => valuesByName[b] - valuesByName[a]);
  const ranks = {};
  let rank = 0, seen = null, shown = 0;
  sorted.forEach(n => {
    shown++;
    if (valuesByName[n] !== seen) { rank = shown; seen = valuesByName[n]; }
    ranks[n] = rank;
  });
  return ranks;
}

/**
 * Rounds a list of {name, amount} payouts to whole cents without losing or inventing money.
 * Any residual from rounding is pushed onto the largest payout, so each pool's dollars paid
 * always equals the dollars collected.
 *
 * @param {Array<Object>} payouts Objects carrying an "amount" property.
 * @param {number} pot The exact pot being distributed.
 * @returns {Array<Object>} The same objects with amounts rounded to cents.
 */
function roundPayoutsToCents(payouts, pot) {
  const list = payouts || [];
  if (!list.length) return list;
  let running = 0;
  list.forEach(p => { p.amount = Math.round((Number(p.amount) || 0) * 100) / 100; running += p.amount; });
  const residual = Math.round((pot - running) * 100) / 100;
  if (residual !== 0) {
    let biggest = 0;
    list.forEach((p, i) => { if (p.amount > list[biggest].amount) biggest = i; });
    list[biggest].amount = Math.round((list[biggest].amount + residual) * 100) / 100;
  }
  return list;
}

/**
 * Splits a pot across ranked groups of tied players.
 *
 * @param {number} pot Total money for the pool.
 * @param {Array<number>} groupSizes Player counts per finishing group, best group first.
 * @param {Array<number>} ladder Percentage per place.
 * @returns {Array<number>} Per-person amount for each group, aligned to groupSizes.
 */
function distributeTieredPot(pot, groupSizes, ladder) {
  const amounts = [];
  let place = 0;
  (groupSizes || []).forEach(size => {
    const members = Number(size) || 0;
    let percent = 0;
    for (let i = 0; i < members; i++) {
      if (place + i < ladder.length) percent += ladder[place + i];
    }
    amounts.push(members > 0 ? (pot * percent / 100) / members : 0);
    place += members;
  });
  return amounts;
}

/**
 * Splits a pot evenly among winners, or evenly among all entrants when nobody won.
 * Used by both consensus pools, where beating the crowd is pass/fail rather than ranked.
 *
 * @param {number} pot Total money for the pool.
 * @param {number} winners How many members beat the consensus.
 * @param {number} entrants How many members entered.
 * @returns {Object} {each, refunded} -- refunded is true when the pot was returned to everyone.
 */
function distributeOrRefund(pot, winners, entrants) {
  if (winners > 0) return { each: pot / winners, refunded: false };
  if (entrants > 0) return { each: pot / entrants, refunded: true };
  return { each: 0, refunded: false };
}

/**
 * Season-points tiers for one week. Ties share a tier and the next DISTINCT correct
 * count takes the next tier down, so a five-way tie for first leaves everyone on 5
 * and the following group on 4.
 *
 * @param {Object} correctByName Map of member name to correct count for the week.
 * @returns {Object} Map of member name to points earned.
 */
function weeklyPlacementPoints(correctByName) {
  const tiers = [5, 4, 3, 2, 1];
  const scores = Object.keys(correctByName)
    .map(name => Number(correctByName[name]))
    .filter(v => !isNaN(v));
  const distinct = Array.from(new Set(scores)).sort((a, b) => b - a);

  const points = {};
  Object.keys(correctByName).forEach(name => {
    const value = Number(correctByName[name]);
    if (isNaN(value)) { points[name] = 0; return; }
    const tierIndex = distinct.indexOf(value);
    points[name] = (tierIndex > -1 && tierIndex < tiers.length) ? tiers[tierIndex] : 0;
  });
  return points;
}

/**
 * Reads one week's results and ranks members for the weekly money.
 *
 * Scores are computed here from the raw picks, outcomes and bonus row rather than read from
 * the sheet's Picks column, so the two things that make this pool different are explicit:
 *   * the weekly BONUS game is worth its multiplier, and weekly placement ranks on points;
 *   * a game the NFL ties is a freebie worth a point to everybody.
 * Raw correct picks are tracked separately because the consensus pools and Season % Correct
 * ignore the bonus entirely.
 *
 * Ranking is points first, then the two tiebreakers. Members identical on all three share a
 * place. Weeks that are not fully scored return null so money is never paid on a part-played week.
 *
 * @param {number} week Week number.
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {Object|null} {groups, matchups, tieGames, perfect}
 */
function weeklyResults(week, ss) {
  ss = fetchSpreadsheet(ss);

  const namesRange = ss.getRangeByName(`NAMES_${week}`);
  const picksRange = ss.getRangeByName(`${LEAGUE}_PICKS_${week}`);
  const outcomeRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
  if (!namesRange || !picksRange || !outcomeRange) return null;

  const outcomes = outcomeRange.getValues()[0].map(v => (v === null || v === undefined) ? '' : v.toString().trim());
  const matchups = outcomes.length;
  const settled = outcomes.filter(v => v !== '').length;
  if (!matchups || settled < matchups) return null; // week not finished -- pay nothing

  // Bonus multipliers along the bottom of the weekly sheet; absent means every game counts once
  const bonusRange = ss.getRangeByName(`${LEAGUE}_BONUS_${week}`);
  const bonuses = bonusRange
    ? bonusRange.getValues()[0].map(v => { const n = Number(v); return (!n || n < 1) ? 1 : n; })
    : outcomes.map(() => 1);

  const tieGames = outcomes.filter(v => v.toUpperCase() === 'TIE').length;
  const tiePoints = outcomes.reduce((sum, o, i) => sum + (o.toUpperCase() === 'TIE' ? bonuses[i] : 0), 0);

  const names = namesRange.getValues().flat();
  const picks = picksRange.getValues();

  // Tiebreaker differences, recomputed from the submitted guesses and the actual scores
  const tb1 = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}`);
  const tb2 = ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}`);
  const tb1Outcome = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}_OUTCOME`);
  const tb2Outcome = ss.getRangeByName(`${LEAGUE}_TIEBREAKER2_${week}_OUTCOME`);
  const tb1Values = tb1 ? tb1.getValues().flat() : [];
  const tb2Values = tb2 ? tb2.getValues().flat() : [];
  const actual1 = tb1Outcome ? Number(tb1Outcome.getValue()) : NaN;
  const actual2 = tb2Outcome ? Number(tb2Outcome.getValue()) : NaN;
  const diff = (guess, actual) => {
    const g = Number(guess);
    if (guess === '' || guess === null || guess === undefined || isNaN(g) || isNaN(actual)) return Infinity;
    return Math.abs(g - actual);
  };

  const players = [], perfect = [];
  names.forEach((name, i) => {
    if (!name) return;
    const row = picks[i] || [];
    const submitted = row.some(p => p !== '' && p !== null && p !== undefined);
    if (!submitted) return;                       // no entry that week

    let points = tiePoints, correct = tieGames;    // NFL ties are a freebie for everyone
    outcomes.forEach((outcome, g) => {
      if (!outcome || outcome.toUpperCase() === 'TIE') return;
      const pick = (row[g] === null || row[g] === undefined) ? '' : row[g].toString().trim();
      if (pick && pick === outcome) { points += bonuses[g]; correct += 1; }
    });

    players.push({
      name: name.toString(),
      points: points,
      correct: correct,
      diff1: diff(tb1Values[i], actual1),
      diff2: diff(tb2Values[i], actual2)
    });
    if (correct === matchups) perfect.push(name.toString());
  });

  // Weekly money and season points both rank on bonus-inclusive points
  players.sort((a, b) => (b.points - a.points) || (a.diff1 - b.diff1) || (a.diff2 - b.diff2));

  const groups = [];
  players.forEach(p => {
    const last = groups[groups.length - 1];
    if (last && last.points === p.points && last.diff1 === p.diff1 && last.diff2 === p.diff2) {
      last.names.push(p.name);
    } else {
      groups.push({ names: [p.name], points: p.points, correct: p.correct, diff1: p.diff1, diff2: p.diff2 });
    }
  });

  return { groups: groups, matchups: matchups, tieGames: tieGames, perfect: perfect };
}

/**
 * Games that count for a week, or 0 when the week is not fully scored yet.
 * Everyone shares the same denominator for a given week.
 *
 * @param {number} week Week number.
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {number} Game count, or 0 if the week is incomplete or missing.
 */
function getGamesPerWeek(week, ss) {
  ss = fetchSpreadsheet(ss);
  const outcomeRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
  if (!outcomeRange) return 0;
  const outcomes = outcomeRange.getValues()[0].map(v => (v === null || v === undefined) ? '' : v.toString().trim());
  const total = outcomes.length;
  const settled = outcomes.filter(v => v !== '').length;
  return (total > 0 && settled === total) ? total : 0;  // an unfinished week is not selectable
}

/**
 * Picks the weeks that count for Pool 3.
 *
 * Sorted by percentage, then by game count, then by week number. Sorting on game count
 * before slicing is what implements the commissioner's rule that a same-percentage week
 * with MORE games wins the last slot -- the higher-denominator week is simply already
 * above the cut, so no separate bubble swap is needed.
 *
 * @param {Array<Object>} weekRecords [{week, correct, games, pct}, ...]
 * @returns {Array<Object>} At most SEASON_PCT_WEEKS records, best first.
 */
function computeBest16Weeks(weekRecords) {
  const sorted = (weekRecords || []).slice().sort((a, b) => {
    if (b.pct !== a.pct) return b.pct - a.pct;      // better percentage first
    if (b.games !== a.games) return b.games - a.games; // same percentage: more games wins
    return a.week - b.week;                          // still tied: earliest week, for determinism
  });
  return sorted.length <= SEASON_PCT_WEEKS ? sorted : sorted.slice(0, SEASON_PCT_WEEKS);
}

/**
 * Pools the chosen weeks into one score. Y is pooled correct over pooled games -- NOT the
 * average of the weekly percentages, which would weight a 12-game week the same as a 16-game one.
 *
 * @param {Array<Object>} weekRecords Every eligible week for one player.
 * @returns {Object} {chosenWeeks, W, X, Y}
 */
function seasonBest16Score(weekRecords) {
  const chosen = computeBest16Weeks(weekRecords);
  const W = chosen.reduce((sum, r) => sum + r.correct, 0);
  const X = chosen.reduce((sum, r) => sum + r.games, 0);
  return { chosenWeeks: chosen, W: W, X: X, Y: X > 0 ? W / X : 0 };
}

/**
 * Pot for one pool. Weekly pools return the pot for a SINGLE week.
 *
 * @param {string} key Pool key from POOLS.
 * @param {number} entries Paid entries.
 * @returns {number} Dollar amount.
 */
function poolPot(key, entries) {
  const pool = POOLS.find(p => p.key === key);
  if (!pool) return 0;
  const count = Number(entries) || 0;
  return pool.cadence === 'weekly' ? pool.fee * count : pool.fee * pool.weeks * count;
}

/**
 * Everything the payout tabs need, computed once from the weekly sheets.
 *
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {Object} Rosters, per-week results, and per-pool money.
 */
function computePayouts(ss) {
  ss = fetchSpreadsheet(ss);
  const memberData = fetchProperties('members') || { memberOrder: [], members: {} };
  const teams = memberData.memberOrder.map(id => memberData.members[id]);
  const names = teams.map(m => (m && m.name) ? m.name.toString() : '').filter(Boolean);
  const entries = names.length;

  const regularWeeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const ladder = payoutLadder(entries);
  const weeklyPot = poolPot('weeklyWins', entries);
  const consensusPot = poolPot('weeklyConsensus', entries);

  const blank = () => { const o = {}; names.forEach(n => o[n] = 0); return o; };
  const result = {
    names: names, entries: entries, ladder: ladder,
    weeklyPot: weeklyPot, consensusPot: consensusPot,
    weeklyMoney: {}, consensusMoney: {}, weeklyPoints: {}, weeklyPlaces: {}, perfectWeekEntries: {},
    totals: { weeklyWins: blank(), weeklyConsensus: blank(), seasonPoints: blank(), perfectWeek: blank(),
              seasonConsensus: blank(), postSeason: blank(), seasonPct: blank() },
    seasonPointsTotal: blank(),
    seasonCorrectTotal: blank(),
    weeklyCorrect: {}, gamesPerWeek: {}, seasonPct: {},
    weeklyCorrectRank: {}, seasonRunningCorrect: {}, seasonCorrectRank: {},
    playoffPoints: {}, playoffPointsTotal: blank(),
    seasonConsensusTotal: 0, seasonConsensusWinners: [],
    perfectWeekWinners: [],
    weeksScored: []
  };

  regularWeeks.forEach(week => {
    const week_ = weeklyResults(week, ss);
    if (!week_) return;                       // week not finished: no money, no points
    result.weeksScored.push(week);

    // --- Weekly Wins: ladder money, ties consuming places below
    const sizes = week_.groups.map(g => g.names.length);
    const perPerson = distributeTieredPot(weeklyPot, sizes, ladder);
    const flat = [];
    let placeCounter = 1;
    const places = {};
    week_.groups.forEach((g, i) => {
      g.names.forEach(n => {
        places[n] = placeCounter;                 // every member of a tied group shares the place
        flat.push({ name: n, amount: perPerson[i] });
      });
      placeCounter += g.names.length;             // ties consume the places below them
    });
    roundPayoutsToCents(flat, weeklyPot);
    const money = {};
    flat.forEach(p => {
      money[p.name] = p.amount;
      result.totals.weeklyWins[p.name] = (result.totals.weeklyWins[p.name] || 0) + p.amount;
    });
    result.weeklyMoney[week] = money;
    result.weeklyPlaces[week] = places;

    // --- Season points: tiers by weekly POINTS (bonus game included), tiebreakers ignored
    const pointsByName = {};
    week_.groups.forEach(g => g.names.forEach(n => pointsByName[n] = g.points));
    const pts = weeklyPlacementPoints(pointsByName);
    result.weeklyPoints[week] = pts;

    // Raw correct picks accumulate separately: the season consensus pool ignores the bonus game
    week_.groups.forEach(g => g.names.forEach(n =>
      result.seasonCorrectTotal[n] = (result.seasonCorrectTotal[n] || 0) + g.correct));

    // Pool 3 needs correct AND games per week, kept apart from the bonus-weighted points
    result.gamesPerWeek[week] = week_.matchups;
    const correctThisWeek = {};
    week_.groups.forEach(g => g.names.forEach(n => correctThisWeek[n] = g.correct));
    result.weeklyCorrect[week] = correctThisWeek;
    const crowd = consensusWeeklyCorrect(week, ss);
    if (crowd) result.seasonConsensusTotal += crowd.adjusted;
    Object.keys(pts).forEach(n => result.seasonPointsTotal[n] = (result.seasonPointsTotal[n] || 0) + pts[n]);

    // --- Weekly Consensus: even split among those who beat the crowd, refund if none
    const flagRange = ss.getRangeByName(`${LEAGUE}_CONSENSUS_${week}`);
    const nameRange = ss.getRangeByName(`NAMES_${week}`);
    if (flagRange && nameRange) {
      const flags = flagRange.getValues().flat();
      const weekNames = nameRange.getValues().flat();
      const winners = [];
      weekNames.forEach((n, i) => { if (n && Number(flags[i]) === 1) winners.push(n.toString()); });
      const split = distributeOrRefund(consensusPot, winners.length, entries);
      const paidTo = split.refunded ? names : winners;
      const cFlat = paidTo.map(n => ({ name: n, amount: split.each }));
      roundPayoutsToCents(cFlat, consensusPot);
      const cMoney = {};
      cFlat.forEach(p => {
        cMoney[p.name] = p.amount;
        result.totals.weeklyConsensus[p.name] = (result.totals.weeklyConsensus[p.name] || 0) + p.amount;
      });
      result.consensusMoney[week] = { money: cMoney, winners: winners.length, refunded: split.refunded };
    }

    // --- Perfect Week: every game correct. Each perfect week is a separate entry into the
    // end-of-season split, so two perfect weeks earns twice the share of one.
    week_.perfect.forEach(n => {
      result.perfectWeekEntries[n] = (result.perfectWeekEntries[n] || 0) + 1;
      if (result.perfectWeekWinners.indexOf(n) === -1) result.perfectWeekWinners.push(n);
    });
  });

  // Every season-long pool keys off the same question: is the regular season finished?
  const pctWeeksComplete = Object.keys(result.weeklyCorrect).length >= REGULAR_SEASON;

  // --- Season Points money: rank by total points, ties consume places below
  const ranked = names.slice().sort((a, b) => result.seasonPointsTotal[b] - result.seasonPointsTotal[a]);
  const pointGroups = [];
  ranked.forEach(n => {
    const last = pointGroups[pointGroups.length - 1];
    if (last && result.seasonPointsTotal[last.names[0]] === result.seasonPointsTotal[n]) last.names.push(n);
    else pointGroups.push({ names: [n] });
  });
  const spPot = poolPot('seasonPoints', entries);
  result.seasonPointsPotGated = true;
  // Season-long pools show standings all the way through but only pay when their period closes,
  // so a dollar figure anywhere in this workbook is money actually won.
  if (pctWeeksComplete) {
    const spEach = distributeTieredPot(spPot, pointGroups.map(g => g.names.length), ladder);
    const spFlat = [];
    pointGroups.forEach((g, i) => g.names.forEach(n => spFlat.push({ name: n, amount: spEach[i] })));
    roundPayoutsToCents(spFlat, spPot);
    spFlat.forEach(p => result.totals.seasonPoints[p.name] = p.amount);
  }
  result.seasonPointsGroups = pointGroups;
  result.seasonPointsPot = spPot;

  // --- Weekly and season rank grids, both on raw correct picks (no tiebreakers, no bonus).
  // Ties share a rank, which is why these can disagree with Wkly Payout -- that pays on
  // bonus-inclusive points with tiebreakers, while these answer "who picked best".
  const rankWeeks = Object.keys(result.weeklyCorrect).map(Number).sort((a, b) => a - b);
  const running = {};
  rankWeeks.forEach(week => {
    result.weeklyCorrectRank[week] = competitionRanks(result.weeklyCorrect[week]);
    const cumulative = {};
    names.forEach(n => {
      const thisWeek = result.weeklyCorrect[week][n];
      running[n] = (running[n] || 0) + ((thisWeek === undefined || thisWeek === null) ? 0 : thisWeek);
      cumulative[n] = running[n];
    });
    result.seasonRunningCorrect[week] = cumulative;
    result.seasonCorrectRank[week] = competitionRanks(cumulative);
  });
  result.rankWeeks = rankWeeks;

  // --- Season % Correct (best 16 weeks). A week a member skipped counts as 0 correct against
  // that week's full denominator, so it lands at 0% and is almost always one of the two dropped.
  const pctWeeks = Object.keys(result.weeklyCorrect).map(Number).sort((a, b) => a - b);
  names.forEach(n => {
    const records = pctWeeks.map(week => {
      const correct = result.weeklyCorrect[week][n];
      const games = result.gamesPerWeek[week];
      const scored = (correct === undefined || correct === null) ? 0 : correct; // missed week = 0
      return { week: week, correct: scored, games: games, pct: games > 0 ? scored / games : 0 };
    }).filter(r => r.games > 0);
    result.seasonPct[n] = seasonBest16Score(records);
  });
  result.seasonPctWeeksAvailable = pctWeeks.length;
  result.regularSeasonComplete = pctWeeks.length >= REGULAR_SEASON;

  const pctPot = poolPot('seasonPct', entries);
  result.seasonPctPot = pctPot;
  // Like the playoff pool, this only pays once its period is over -- until then the tab shows
  // W/X/Y and a provisional rank so members can track where they stand.
  if (result.regularSeasonComplete) {
    const pctRanked = names.slice().sort((a, b) => result.seasonPct[b].Y - result.seasonPct[a].Y);
    const pctGroups = [];
    pctRanked.forEach(n => {
      const last = pctGroups[pctGroups.length - 1];
      if (last && result.seasonPct[last.names[0]].Y === result.seasonPct[n].Y) last.names.push(n);
      else pctGroups.push({ names: [n] });
    });
    const pctEach = distributeTieredPot(pctPot, pctGroups.map(g => g.names.length), ladder);
    const pctFlat = [];
    pctGroups.forEach((g, i) => g.names.forEach(n => pctFlat.push({ name: n, amount: pctEach[i] })));
    roundPayoutsToCents(pctFlat, pctPot);
    pctFlat.forEach(p => result.totals.seasonPct[p.name] = p.amount);
  }

  // One worked example in the log, so the best-16 maths can be checked by hand
  if (names.length && result.seasonPct[names[0]]) {
    const sample = result.seasonPct[names[0]];
    Logger.log(`📊 Best-16 example for "${names[0]}": kept ${sample.chosenWeeks.length} of ${pctWeeks.length} eligible weeks ` +
      `(${sample.chosenWeeks.map(r => `wk${r.week} ${r.correct}/${r.games}`).join(', ')}) ` +
      `-> W=${sample.W}, X=${sample.X}, Y=${(sample.Y * 100).toFixed(2)}%`);
  }

  // --- Season Consensus: beat the crowd's season total outright, on raw correct picks.
  // Even split among winners, refunded to everyone if the crowd wins.
  const scPot = poolPot('seasonConsensus', entries);
  result.seasonConsensusWinners = names.filter(n => (result.seasonCorrectTotal[n] || 0) > result.seasonConsensusTotal);
  const scSplit = distributeOrRefund(scPot, result.seasonConsensusWinners.length, entries);
  if (pctWeeksComplete) {
    const scFlat = (scSplit.refunded ? names : result.seasonConsensusWinners).map(n => ({ name: n, amount: scSplit.each }));
    roundPayoutsToCents(scFlat, scPot);
    scFlat.forEach(p => result.totals.seasonConsensus[p.name] = p.amount);
  }
  result.seasonConsensusPot = scPot;
  result.seasonConsensusRefunded = scSplit.refunded;

  // --- Post Season: same 5/4/3/2/1 tiers as the season points pool, over the playoff weeks only,
  // carried as one cumulative total rather than weekly pots.
  PLAYOFF_WEEKS.forEach(week => {
    const playoff = weeklyResults(week, ss);
    if (!playoff) return;
    const byName = {};
    playoff.groups.forEach(g => g.names.forEach(n => byName[n] = g.points));
    const pts = weeklyPlacementPoints(byName);
    result.playoffPoints[week] = pts;
    Object.keys(pts).forEach(n => result.playoffPointsTotal[n] = (result.playoffPointsTotal[n] || 0) + pts[n]);
  });
  const playoffScored = Object.keys(result.playoffPoints).length;
  const psPot = poolPot('postSeason', entries);
  result.postSeasonPot = psPot;
  result.playoffWeeksScored = playoffScored;
  result.playoffComplete = playoffScored === PLAYOFF_WEEKS.length;
  // Points accumulate visibly all through the playoffs, but the pot is only awarded once
  // every playoff week is scored -- nobody sees dollars they have not won yet.
  if (result.playoffComplete) {
    const psRanked = names.slice().sort((a, b) => result.playoffPointsTotal[b] - result.playoffPointsTotal[a]);
    const psGroups = [];
    psRanked.forEach(n => {
      const last = psGroups[psGroups.length - 1];
      if (last && result.playoffPointsTotal[last.names[0]] === result.playoffPointsTotal[n]) last.names.push(n);
      else psGroups.push({ names: [n] });
    });
    const psEach = distributeTieredPot(psPot, psGroups.map(g => g.names.length), ladder);
    const psFlat = [];
    psGroups.forEach((g, i) => g.names.forEach(n => psFlat.push({ name: n, amount: psEach[i] })));
    roundPayoutsToCents(psFlat, psPot);
    psFlat.forEach(p => result.totals.postSeason[p.name] = p.amount);
  }

  // --- Perfect Week money: shares are proportional to how many perfect weeks a member had,
  // so two perfect weeks against somebody else's one splits the pot 2/3 to 1/3. Nobody perfect
  // all season means the pot is refunded evenly.
  const pwPot = poolPot('perfectWeek', entries);
  const totalEntries = Object.keys(result.perfectWeekEntries)
    .reduce((sum, n) => sum + result.perfectWeekEntries[n], 0);
  const pwFlat = [];
  if (totalEntries > 0) {
    Object.keys(result.perfectWeekEntries).forEach(n =>
      pwFlat.push({ name: n, amount: pwPot * (result.perfectWeekEntries[n] / totalEntries) }));
    result.perfectWeekRefunded = false;
  } else {
    names.forEach(n => pwFlat.push({ name: n, amount: entries ? pwPot / entries : 0 }));
    result.perfectWeekRefunded = true;
  }
  roundPayoutsToCents(pwFlat, pwPot);
  pwFlat.forEach(p => result.totals.perfectWeek[p.name] = p.amount);
  result.perfectWeekPot = pwPot;
  result.perfectWeekShares = totalEntries;

  return result;
}

/** Shared look and feel for the money grids. */
function styleMoneyGrid(sheet, rows, cols, moneyRange) {
  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold');
  sheet.getRange(rows, 1, 1, cols).setBackground('#e6e6e6').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 80);
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  if (moneyRange) moneyRange.setNumberFormat('$#,##0.00;[Red]-$#,##0.00;""');
}

/**
 * WKLY PAYOUT -- dollars won in the main weekly pick 'em, week by week.
 */
function wklyPayoutSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const names = data.names;
  const rows = names.length + 2, cols = weeks.length + 2;

  let sheet = ss.getSheetByName('Wkly Payout') || ss.insertSheet('Wkly Payout');
  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  const header = ['Team', 'Total'].concat(weeks.map(w => `WK${w}`));
  sheet.getRange(1, 1, 1, cols).setValues([header]);

  const body = names.map(n => [n, data.totals.weeklyWins[n] || 0]
    .concat(weeks.map(w => (data.weeklyMoney[w] && data.weeklyMoney[w][n]) ? data.weeklyMoney[w][n] : '')));
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const totalRow = ['Pot', names.reduce((s, n) => s + (data.totals.weeklyWins[n] || 0), 0)]
    .concat(weeks.map(w => data.weeklyMoney[w] ? data.weeklyPot : ''));
  sheet.getRange(rows, 1, 1, cols).setValues([totalRow]);

  styleMoneyGrid(sheet, rows, cols, sheet.getRange(2, 2, rows - 1, cols - 1));
  sheet.getRange(1, 1).setNote(`Dollars won in the main weekly pick 'ems. ${data.ladder.length} places paid at ${data.ladder.join('/')}% of each week's $${data.weeklyPot.toFixed(2)} pot. Tied members take the places below them and split the combined money. Blank means the week is not fully scored yet. Refreshed by "Update Payouts".`);
  return sheet;
}

/**
 * WKLY CONSENSUS -- dollars won for beating the crowd, week by week.
 */
function wklyConsensusPayoutSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const names = data.names;
  const rows = names.length + 3, cols = weeks.length + 2;

  let sheet = ss.getSheetByName('Wkly Consensus') || ss.insertSheet('Wkly Consensus');
  sheet.clear();
  sheet.setTabColor(generalTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  sheet.getRange(1, 1, 1, cols).setValues([['Team', 'Total'].concat(weeks.map(w => `WK${w}`))]);

  const body = names.map(n => [n, data.totals.weeklyConsensus[n] || 0]
    .concat(weeks.map(w => (data.consensusMoney[w] && data.consensusMoney[w].money[n]) ? data.consensusMoney[w].money[n] : '')));
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const winnersRow = ['Winners', ''].concat(weeks.map(w => data.consensusMoney[w]
    ? (data.consensusMoney[w].refunded ? 'refund' : data.consensusMoney[w].winners) : ''));
  sheet.getRange(rows - 1, 1, 1, cols).setValues([winnersRow]);

  const potRow = ['Pot', names.reduce((s, n) => s + (data.totals.weeklyConsensus[n] || 0), 0)]
    .concat(weeks.map(w => data.consensusMoney[w] ? data.consensusPot : ''));
  sheet.getRange(rows, 1, 1, cols).setValues([potRow]);

  styleMoneyGrid(sheet, rows, cols, sheet.getRange(2, 2, names.length, cols - 1));
  sheet.getRange(rows - 1, 1, 1, cols).setBackground('#f3f3f3');
  sheet.getRange(1, 1).setNote(`Dollars won for finishing with strictly more correct picks than the group's consensus row. Each week's $${data.consensusPot.toFixed(2)} pot is split evenly among the winners; if nobody beats the consensus the pot is refunded evenly to every entrant and the Winners row reads "refund". Refreshed by "Update Payouts".`);
  return sheet;
}

/**
 * SEASON POINTS -- weekly 5/4/3/2/1 tiers, season total, rank, and money.
 */
function seasonPointsSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const names = data.names.slice().sort((a, b) => data.seasonPointsTotal[b] - data.seasonPointsTotal[a]);
  const leader = names.length ? data.seasonPointsTotal[names[0]] : 0;
  const rows = names.length + 1, cols = weeks.length + 5;

  let sheet = ss.getSheetByName('Season Points') || ss.insertSheet('Season Points');
  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  sheet.getRange(1, 1, 1, cols).setValues([['Team', 'Rank', '$ Won', 'Off Lead', 'Total Points']
    .concat(weeks.map(w => `WK${w}`))]);

  // Rank shares a number across ties, and the next distinct total skips ahead
  let rank = 0, seen = null, shown = 0;
  const body = names.map(n => {
    shown++;
    if (data.seasonPointsTotal[n] !== seen) { rank = shown; seen = data.seasonPointsTotal[n]; }
    return [n, rank, data.totals.seasonPoints[n] || 0,
            data.seasonPointsTotal[n] - leader === 0 ? '' : data.seasonPointsTotal[n] - leader,
            data.seasonPointsTotal[n]]
      .concat(weeks.map(w => (data.weeklyPoints[w] && data.weeklyPoints[w][n] !== undefined) ? data.weeklyPoints[w][n] : ''));
  });
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold');
  sheet.getRange(2, 3, Math.max(names.length, 1), 1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00;""');
  sheet.setColumnWidth(1, 150);
  sheet.setFrozenColumns(5);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setNote(`Weekly tiers of 5/4/3/2/1 by correct picks only -- tiebreakers are not used here. Tied members all take that tier's full points and the next distinct score drops a tier. Season pot of $${data.seasonPointsPot.toFixed(2)} pays ${data.ladder.length} places at ${data.ladder.join('/')}%, with ties consuming the places below. Points show all season; the money is only assigned once all ${REGULAR_SEASON} regular-season weeks are scored. Refreshed by "Update Payouts".`);
  return sheet;
}

/**
 * SUMMARY PAYOUT -- every member's money across every pool, plus the fee schedule
 * block that defines each pot. This tab is the one members read to see where they stand.
 */
function summaryPayoutSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const memberData = fetchProperties('members') || { memberOrder: [], members: {} };
  const byName = {};
  memberData.memberOrder.forEach(id => {
    const m = memberData.members[id];
    if (m && m.name) byName[m.name.toString()] = m;
  });

  const poolCols = POOLS.map(p => p.label);
  const header = ['Team', 'Manager', 'Won'].concat(poolCols).concat(['Entry', 'Fines', 'Net', 'Paid']);
  const names = data.names;
  const firstPool = 4;                       // column D
  const cols = header.length;
  const configTop = names.length + 3;        // one blank row after the roster
  const rows = configTop + 7;

  let sheet = ss.getSheetByName('Summary Payout') || ss.insertSheet('Summary Payout');
  sheet.clear();
  sheet.setTabColor(configTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);
  sheet.getRange(1, 1, 1, cols).setValues([header]);

  const entryTotal = POOLS.reduce((s, p) => s + p.fee * p.weeks, 0);
  const body = names.map(n => {
    const m = byName[n] || {};
    const won = POOLS.map(p => {
      if (p.key === 'weeklyWins') return data.totals.weeklyWins[n] || 0;
      if (p.key === 'weeklyConsensus') return data.totals.weeklyConsensus[n] || 0;
      if (p.key === 'seasonPoints') return data.totals.seasonPoints[n] || 0;
      if (p.key === 'perfectWeek') return data.totals.perfectWeek[n] || 0;
      if (p.key === 'seasonConsensus') return data.totals.seasonConsensus[n] || 0;
      if (p.key === 'postSeason') return data.totals.postSeason[n] || 0;
      if (p.key === 'seasonPct') return data.totals.seasonPct[n] || 0;
      return '';
    });
    return [n, m.manager || ''].concat([''], won, [entryTotal, '', '', m.paid ? 'YES' : '']);
  });
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  // Won = sum of the pool columns; Net = Won - Entry - Fines
  names.forEach((n, i) => {
    const r = i + 2;
    sheet.getRange(r, 3).setFormula(`=IFERROR(SUM(${sheet.getRange(r, firstPool).getA1Notation()}:${sheet.getRange(r, firstPool + POOLS.length - 1).getA1Notation()}),0)`);
    const entryCell = sheet.getRange(r, firstPool + POOLS.length).getA1Notation();
    const finesCell = sheet.getRange(r, firstPool + POOLS.length + 1).getA1Notation();
    sheet.getRange(r, firstPool + POOLS.length + 2)
      .setFormula(`=IFERROR(C${r}-${entryCell}-N(${finesCell}),)`);
  });

  // Fee schedule block -- the definition of every pot, mirroring the 2025 layout
  const labels = ['cost per entry per week', 'Weeks', 'Total per entry', 'Entries', 'POT', '% of total pot'];
  const totalPot = POOLS.reduce((s, p) => s + p.fee * p.weeks * data.entries, 0);
  labels.forEach((label, i) => {
    const r = configTop + i;
    sheet.getRange(r, 2).setValue(label);
    POOLS.forEach((p, c) => {
      const col = firstPool + c;
      let v = '';
      if (i === 0) v = p.fee;
      if (i === 1) v = p.weeks;
      if (i === 2) v = p.fee * p.weeks;
      if (i === 3) v = data.entries;
      if (i === 4) v = p.fee * p.weeks * data.entries;
      if (i === 5) v = totalPot ? (p.fee * p.weeks * data.entries) / totalPot : 0;
      sheet.getRange(r, col).setValue(v);
    });
  });
  sheet.getRange(configTop + 2, 3).setValue(entryTotal);
  sheet.getRange(configTop + 4, 3).setValue(totalPot);
  sheet.getRange(configTop + 5, firstPool, 1, POOLS.length).setNumberFormat('0%');
  sheet.getRange(configTop, firstPool, 1, POOLS.length).setNumberFormat('$#,##0.00');
  sheet.getRange(configTop + 2, firstPool, 1, POOLS.length).setNumberFormat('$#,##0.00');
  sheet.getRange(configTop + 4, firstPool, 1, POOLS.length).setNumberFormat('$#,##0.00');
  sheet.getRange(configTop + 2, 3).setNumberFormat('$#,##0.00');
  sheet.getRange(configTop + 4, 3).setNumberFormat('$#,##0.00');

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 2).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 44);
  if (names.length) {
    sheet.getRange(2, 3, names.length, POOLS.length + 3).setNumberFormat('$#,##0.00;[Red]-$#,##0.00;""');
    sheet.getRange(2, 3, names.length, 1).setFontWeight('bold');
  }
  sheet.getRange(configTop, 1, 6, cols).setBackground('#f3f3f3');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 130);
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  sheet.getRange(1, firstPool + POOLS.length + 1).setNote(`Manual column. Type any fine owed and it is subtracted from that member's Net.`);
  sheet.getRange(1, 3).setNote(`Total won across every pool. Net subtracts the entry fee and any fines.`);
  sheet.getRange(configTop, 2).setNote(`This block defines every pot: fee per entry per week x weeks x entries. Change a fee here and re-run "Update Payouts" to push it through every tab.`);
  return sheet;
}

/**
 * Writes each member's finishing place into the Place column of every scored weekly sheet,
 * so the WK sheet and Wkly Payout tell the same story. Place 1 still means first, which keeps
 * the WINNERS tab (which joins names where the place equals 1) working unchanged.
 */
function writeWeeklyPlaces(ss, data) {
  Object.keys(data.weeklyPlaces).forEach(week => {
    const nameRange = ss.getRangeByName(`NAMES_${week}`);
    const placeRange = ss.getRangeByName(`WIN_${week}`);
    if (!nameRange || !placeRange) return;
    const names = nameRange.getValues().flat();
    const places = data.weeklyPlaces[week];
    placeRange.setValues(names.map(n => {
      const key = n ? n.toString() : '';
      return [(key && places[key]) ? places[key] : ''];
    }));
  });
}

/**
 * Builds the payout tabs without any dialogs, for use during initial setup.
 */
function updatePayoutSheetsQuietly(ss) {
  ss = fetchSpreadsheet(ss);
  try {
    const data = computePayouts(ss);
    if (!data.entries) return;
    writeWeeklyPlaces(ss, data);
    wklyPayoutSheet(ss, data);
    wklyConsensusPayoutSheet(ss, data);
    seasonPointsSheet(ss, data);
    seasonTotalsSheet(ss, data);
    if (data.playoffWeeksScored) playoffPointsSheet(ss, data);
    summaryPayoutSheet(ss, data);
    standingsSheet(ss, data);
    wklyRankSheet(ss, data);
    seasonRankSheet(ss, data);
    rankJumpChartSheet(ss, data);
  } catch (err) {
    Logger.log(`⚠️ Could not deploy payout sheets: ${err.stack}`);
  }
}

/**
 * PLAYOFF POINTS -- the Post Season pool: 5/4/3/2/1 per playoff week, one cumulative total.
 */
function playoffPointsSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const names = data.names.slice().sort((a, b) => data.playoffPointsTotal[b] - data.playoffPointsTotal[a]);
  const rows = names.length + 1, cols = PLAYOFF_WEEKS.length + 3;

  let sheet = ss.getSheetByName('Playoff Points') || ss.insertSheet('Playoff Points');
  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  const roundName = { 19: 'Wild Card', 20: 'Divisional', 21: 'Conference', 23: 'Super Bowl' };
  sheet.getRange(1, 1, 1, cols).setValues([['Team', 'Rank', 'Total Points']
    .concat(PLAYOFF_WEEKS.map(w => roundName[w] || `WK${w}`))]);

  let rank = 0, seen = null, shown = 0;
  const body = names.map(n => {
    shown++;
    if (data.playoffPointsTotal[n] !== seen) { rank = shown; seen = data.playoffPointsTotal[n]; }
    return [n, rank, data.playoffPointsTotal[n] || 0]
      .concat(PLAYOFF_WEEKS.map(w => (data.playoffPoints[w] && data.playoffPoints[w][n] !== undefined) ? data.playoffPoints[w][n] : ''));
  });
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold').setWrap(true);
  if (names.length) sheet.getRange(2, 3, names.length, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setFrozenColumns(3);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setNote(`Separate entry from the regular season. Each playoff week awards 5/4/3/2/1 by points and ties share a tier, so this is a running total you can track as the playoffs go. Money is deliberately not shown here -- the $${data.postSeasonPot.toFixed(2)} pot pays ${data.ladder.length} places at ${data.ladder.join('/')}% only once all ${PLAYOFF_WEEKS.length} playoff weeks are scored, and it appears on Summary Payout and Standings then. ${data.playoffWeeksScored} of ${PLAYOFF_WEEKS.length} playoff weeks scored so far.`);
  return sheet;
}

/**
 * Shared builder for the two rank grids. Both are members x weeks with a rank in each cell.
 */
function buildRankGrid(ss, tabName, data, rankByWeek, currentValueByName, noteText) {
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const latest = data.rankWeeks.length ? data.rankWeeks[data.rankWeeks.length - 1] : null;
  const finalRank = latest ? rankByWeek[latest] : {};
  const names = data.names.slice().sort((a, b) =>
    (finalRank[a] || 999) - (finalRank[b] || 999) || a.localeCompare(b));
  const rows = names.length + 1, cols = weeks.length + 2;

  let sheet = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sheet.clear();
  sheet.setTabColor(generalTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  sheet.getRange(1, 1, 1, cols).setValues([['Team', 'Now'].concat(weeks.map(w => `WK${w}`))]);
  const body = names.map(n => [n, (finalRank[n] === undefined ? '' : finalRank[n])]
    .concat(weeks.map(w => (rankByWeek[w] && rankByWeek[w][n] !== undefined) ? rankByWeek[w][n] : '')));
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidths(3, weeks.length, 34);
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // Rank 1 is best, so the gradient runs dark-to-light from the top
  if (names.length) {
    sheet.clearConditionalFormatRules();
    sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, String(names.length))
      .setGradientMinpointWithValue('#5EDCFF', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setRanges([sheet.getRange(2, 2, names.length, cols - 1)])
      .build()]);
  }
  sheet.getRange(1, 1).setNote(noteText);
  return { sheet: sheet, names: names, rows: rows, cols: cols };
}

/**
 * WKLY RANK -- where each member finished each week on raw correct picks alone.
 */
function wklyRankSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const latest = data.rankWeeks.length ? data.rankWeeks[data.rankWeeks.length - 1] : null;
  return buildRankGrid(ss, 'Wkly Rank', data, data.weeklyCorrectRank,
    latest ? data.weeklyCorrectRank[latest] : {},
    `Each week's finishing order by correct picks only -- no tiebreakers and no bonus game, so ties share a rank. This answers "who picked best". It can differ from Wkly Payout, which pays on bonus-inclusive points with the tiebreakers applied. "Now" is the most recent scored week.`).sheet;
}

/**
 * SEASON RANK -- running position by cumulative correct picks after each week.
 */
function seasonRankSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  return buildRankGrid(ss, 'Season Rank', data, data.seasonCorrectRank, {},
    `Running rank by total correct picks through each week. Ties share a rank. A week someone missed adds nothing, so skipping a week costs ground here. Bonus games are not counted.`).sheet;
}

/**
 * RANK JUMP CHART -- season rank week by week, the change from the previous week, and a
 * line chart of everyone's climb and slide across the season.
 */
function rankJumpChartSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const ranks = data.seasonCorrectRank;
  const latest = data.rankWeeks.length ? data.rankWeeks[data.rankWeeks.length - 1] : null;
  const previous = data.rankWeeks.length > 1 ? data.rankWeeks[data.rankWeeks.length - 2] : null;
  const finalRank = latest ? ranks[latest] : {};
  const names = data.names.slice().sort((a, b) =>
    (finalRank[a] || 999) - (finalRank[b] || 999) || a.localeCompare(b));

  const cols = weeks.length + 3;               // Team + 18 weeks + Rank + Change
  const rows = names.length + 1;

  let sheet = ss.getSheetByName('Rank Jump Chart') || ss.insertSheet('Rank Jump Chart');
  sheet.getCharts().forEach(c => sheet.removeChart(c));   // rebuild cleanly
  sheet.clear();
  sheet.setTabColor(generalTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  sheet.getRange(1, 1, 1, cols).setValues([['Team'].concat(weeks.map(w => `WK${w}`)).concat(['Rank', 'Change'])]);
  const body = names.map(n => {
    const now = finalRank[n], before = previous ? ranks[previous][n] : undefined;
    let change = '';
    if (now !== undefined && before !== undefined) {
      const delta = before - now;                          // positive means climbed
      change = delta === 0 ? '--' : (delta > 0 ? `\u25b2${delta}` : `\u25bc${Math.abs(delta)}`);
    }
    return [n].concat(weeks.map(w => (ranks[w] && ranks[w][n] !== undefined) ? ranks[w][n] : ''))
      .concat([now === undefined ? '' : now, change]);
  });
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidths(2, weeks.length, 34);
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);

  if (names.length) {
    // Climbers green, sliders red, in the Change column
    sheet.clearConditionalFormatRules();
    const changeRange = sheet.getRange(2, cols, names.length, 1);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('\u25b2')
        .setFontColor('#188038').setBold(true).setRanges([changeRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('\u25bc')
        .setFontColor('#c5221f').setRanges([changeRange]).build()
    ]);

    // Rank 1 belongs at the TOP of the chart, so the vertical axis is inverted
    const chart = sheet.newChart().asLineChart()
      .addRange(sheet.getRange(1, 1, rows, weeks.length + 1))
      .setTransposeRowsAndColumns(true)
      .setOption('title', `Season rank by week -- lower is better`)
      .setOption('vAxis', { direction: -1, title: 'Rank', gridlines: { count: -1 } })
      .setOption('hAxis', { title: 'Week' })
      .setOption('legend', { position: 'right' })
      .setOption('height', 420)
      .setOption('width', 900)
      .setPosition(rows + 3, 1, 0, 0)
      .build();
    sheet.insertChart(chart);
  }

  sheet.getRange(1, 1).setNote(`Season rank week by week, on cumulative correct picks. "Change" compares the latest scored week with the one before it -- green climbed, red slid. The chart plots every member's rank across the season with the axis inverted, so first place sits at the top.`);
  return sheet;
}

/**
 * SEASON TOTALS -- Pool 3. Weekly percentages across the regular season, with each member's
 * best 16 pooled into W, X and Y. Weeks a member DROPPED are greyed out, so anyone can see
 * which two weeks were discarded and check the maths by hand.
 */
function seasonTotalsSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);
  const weeks = Array.from({ length: REGULAR_SEASON }, (_, i) => i + 1);
  const names = data.names.slice().sort((a, b) => data.seasonPct[b].Y - data.seasonPct[a].Y);
  const rows = names.length + 1, cols = weeks.length + 6;

  let sheet = ss.getSheetByName('Season Totals') || ss.insertSheet('Season Totals');
  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  adjustRows(sheet, rows); adjustColumns(sheet, cols);

  sheet.getRange(1, 1, 1, cols).setValues([['Team', 'Rank', 'W', 'X', 'Y', 'Weeks Kept']
    .concat(weeks.map(w => `WK${w}`))]);

  let rank = 0, seen = null, shown = 0;
  const body = [], fontColors = [];
  names.forEach(n => {
    shown++;
    const score = data.seasonPct[n];
    if (score.Y !== seen) { rank = shown; seen = score.Y; }
    const kept = {};
    score.chosenWeeks.forEach(r => kept[r.week] = true);
    body.push([n, rank, score.W, score.X, score.Y, score.chosenWeeks.length]
      .concat(weeks.map(w => {
        const games = data.gamesPerWeek[w];
        if (!games) return '';                                   // week not scored yet
        const correct = data.weeklyCorrect[w][n];
        return ((correct === undefined || correct === null) ? 0 : correct) / games;
      })));
    // Dropped weeks are faded rather than hidden -- the number still matters for checking
    fontColors.push(['#000000', '#000000', '#000000', '#000000', '#000000', '#666666']
      .concat(weeks.map(w => (!data.gamesPerWeek[w] || kept[w]) ? '#000000' : '#c0c0c0')));
  });
  if (body.length) {
    sheet.getRange(2, 1, body.length, cols).setValues(body);
    sheet.getRange(2, 1, body.length, cols).setFontColors(fontColors);
    sheet.getRange(2, 5, body.length, 1).setNumberFormat('##0.00%').setFontWeight('bold');
    sheet.getRange(2, 7, body.length, weeks.length).setNumberFormat('##0.0%;;""');
  }

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 1, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 40);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidths(2, 5, 55);
  sheet.setFrozenColumns(6);
  sheet.setFrozenRows(1);

  sheet.getRange(1, 3).setNote(`W: correct picks across the member's best ${SEASON_PCT_WEEKS} weeks.`);
  sheet.getRange(1, 4).setNote(`X: games in those same weeks. Different members keep different weeks, so X varies.`);
  sheet.getRange(1, 5).setNote(`Y = W / X, the ranking value for this pool. It is pooled correct over pooled games, NOT the average of the weekly percentages.`);
  sheet.getRange(1, 1).setNote(`Season % Correct, best ${SEASON_PCT_WEEKS} of ${REGULAR_SEASON} weeks. Faded percentages are weeks that member dropped. A week someone skipped counts as 0 correct against the full game count, so it lands at 0% and is almost always dropped. Bonus games do not count here -- this is straight picks correct. ${data.regularSeasonComplete ? `The $${data.seasonPctPot.toFixed(2)} pot pays ${data.ladder.length} places at ${data.ladder.join('/')}%.` : `Provisional: ${data.seasonPctWeeksAvailable} of ${REGULAR_SEASON} weeks scored, so no money is assigned yet.`}`);
  return sheet;
}

/**
 * STANDINGS -- the member-facing view. One row per team, money to date, and where they
 * sit in each pool. Everything here is derived; nothing is entered by hand.
 */
function standingsSheet(ss, data) {
  ss = fetchSpreadsheet(ss);
  data = data || computePayouts(ss);

  const totalWon = {};
  data.names.forEach(n => {
    totalWon[n] = (data.totals.weeklyWins[n] || 0) + (data.totals.weeklyConsensus[n] || 0)
      + (data.totals.seasonPoints[n] || 0) + (data.totals.seasonConsensus[n] || 0)
      + (data.totals.postSeason[n] || 0) + (data.totals.perfectWeek[n] || 0)
      + (data.totals.seasonPct[n] || 0);
  });
  const names = data.names.slice().sort((a, b) => totalWon[b] - totalWon[a]);

  const header = ['Rank', 'Team', 'Won', 'Weekly Wins', 'Weeks Won', 'Consensus $', 'Weeks Beat Crowd',
                  'Season Points', 'Correct Picks', 'Best-16 %', 'Perfect Weeks', 'Playoff Points'];
  const rows = names.length + 1, cols = header.length;

  let sheet = ss.getSheetByName('Standings') || ss.insertSheet('Standings', 0);
  sheet.clear();
  sheet.setTabColor('#34a853');
  adjustRows(sheet, rows); adjustColumns(sheet, cols);
  sheet.getRange(1, 1, 1, cols).setValues([header]);

  let rank = 0, seen = null, shown = 0;
  const body = names.map(n => {
    shown++;
    if (totalWon[n] !== seen) { rank = shown; seen = totalWon[n]; }
    const weeksWon = Object.keys(data.weeklyPlaces).filter(w => data.weeklyPlaces[w][n] === 1).length;
    const weeksBeat = Object.keys(data.consensusMoney)
      .filter(w => !data.consensusMoney[w].refunded && data.consensusMoney[w].money[n]).length;
    return [rank, n, totalWon[n],
            data.totals.weeklyWins[n] || 0, weeksWon,
            data.totals.weeklyConsensus[n] || 0, weeksBeat,
            data.seasonPointsTotal[n] || 0, data.seasonCorrectTotal[n] || 0,
            (data.seasonPct[n] ? data.seasonPct[n].Y : 0),
            data.perfectWeekEntries[n] || 0, data.playoffPointsTotal[n] || 0];
  });
  if (body.length) sheet.getRange(2, 1, body.length, cols).setValues(body);

  const all = sheet.getRange(1, 1, rows, cols);
  all.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Montserrat').setFontSize(10);
  sheet.getRange(1, 2, rows, 1).setHorizontalAlignment('left');
  sheet.getRange(1, 1, 1, cols).setBackground('black').setFontColor('white').setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 42);
  if (names.length) {
    sheet.getRange(2, 3, names.length, 1).setNumberFormat('$#,##0.00').setFontWeight('bold');
    sheet.getRange(2, 4, names.length, 1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00;""');
    sheet.getRange(2, 6, names.length, 1).setNumberFormat('$#,##0.00;[Red]-$#,##0.00;""');
    sheet.getRange(2, 10, names.length, 1).setNumberFormat('##0.0%;;""');
    const moneyRange = sheet.getRange(2, 3, names.length, 1);
    sheet.clearConditionalFormatRules();
    sheet.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#75F0A1').setGradientMinpoint('#FFFFFF')
      .setRanges([moneyRange]).build()]);
  }
  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 160);
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1).setNote(`Where everyone stands, sorted by money won. Weekly figures only count weeks that are fully scored. Playoff Points is a running points total -- that pool's money is not paid until all four playoff weeks are in. Refreshed by "Update Payouts" or "Close Out Week".`);
  return sheet;
}

/**
 * Menu action: recompute every payout tab from the weekly sheets.
 */
function updatePayouts() {
  const ss = fetchSpreadsheet();
  const ui = fetchUi();
  try {
    ss.toast(`Reading weekly results...`, `💵 UPDATING PAYOUTS`);
    const data = computePayouts(ss);
    if (!data.entries) {
      ui.alert(`💵 NO MEMBERS`, `There are no members yet, so there is nothing to pay out. Add members or import sign-ups first.`, ui.ButtonSet.OK);
      return;
    }
    writeWeeklyPlaces(ss, data);
    wklyPayoutSheet(ss, data);
    wklyConsensusPayoutSheet(ss, data);
    seasonPointsSheet(ss, data);
    seasonTotalsSheet(ss, data);
    playoffPointsSheet(ss, data);
    summaryPayoutSheet(ss, data);
    standingsSheet(ss, data);
    wklyRankSheet(ss, data);
    seasonRankSheet(ss, data);
    rankJumpChartSheet(ss, data);

    const scored = data.weeksScored.length;
    const perfect = data.perfectWeekWinners.length;
    ss.toast(`Payout tabs updated`, `✅ DONE`);
    ui.alert(`✅ PAYOUTS UPDATED`,
      `${data.entries} entries, ${data.ladder.length} places paid at ${data.ladder.join('/')}%.\n\n` +
      `Weeks fully scored: ${scored}\n` +
      `Weekly pot: $${data.weeklyPot.toFixed(2)}   Consensus pot: $${data.consensusPot.toFixed(2)}\n` +
      `Perfect weeks: ${perfect ? data.perfectWeekWinners.join(', ') : 'none yet' + (data.perfectWeekRefunded ? ' (pot would refund evenly)' : '')}\n\n` +
      `Season % Correct, Season Consensus and Post Season columns stay blank until those tabs are built.`,
      ui.ButtonSet.OK);
  } catch (err) {
    Logger.log(`❌ updatePayouts failed: ${err.stack}`);
    ui.alert(`⚠️ PAYOUT ERROR`, `Could not update the payout tabs:\n\n${err.message}`, ui.ButtonSet.OK);
  }
}

// ============================================================================================================================================
// SIGN-UP FORM
// ============================================================================================================================================
//
// A standalone form for collecting entries before the season starts. Ben's tool only
// enrolls people through a weekly form's "New User" page, which means signing up
// requires submitting Week 1 picks -- no use in August. This form stands on its own,
// and "Import Sign-Ups" turns the responses into members so the roster is already
// populated by the time the Week 1 form gets built.

const SIGNUP_Q_TEAM = 'Team Name';
const SIGNUP_Q_MANAGER = 'Your Name';
const SIGNUP_Q_EMAIL = 'Email';
const SIGNUP_Q_CONFIRM = "Confirm your entry";

/**
 * Menu entry point. Creates the sign-up form on first use, then hands back the link.
 */
function launchSignupForm() {
  const ui = fetchUi();
  const ss = fetchSpreadsheet();

  let signup = fetchProperties('signup');
  let form = null;

  if (signup && signup.formId) {
    try {
      const file = DriveApp.getFileById(signup.formId);
      if (file && !file.isTrashed()) form = FormApp.openById(signup.formId);
    } catch (err) {
      Logger.log(`⚠️ Stored sign-up form could not be opened: ${err.message}`);
    }
    if (!form) {
      Logger.log(`🧹 Clearing stale sign-up form reference`);
      deleteProperties('signup');
    }
  }

  if (!form) {
    const answer = ui.alert(`✍️ CREATE SIGN-UP FORM`,
      `No sign-up form exists yet.\n\nCreate one now? It collects a team name, the member's name, and an email address, and it does not ask for any picks -- so you can send it out well before week 1.`,
      ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) {
      ss.toast(`Sign-up form creation canceled`,`⛔ CANCELED`);
      return;
    }
    form = createSignupForm(ss);
    if (!form) return;
  }

  const signupData = fetchProperties('signup') || {};
  const url = signupData.publishedUrl || form.getPublishedUrl();
  showLinkDialog(url, `✍️ Sign-Up Form`, `Sign-Up Form`,
    `\nShare this link with your group now. Run "Import Sign-Ups" from the menu whenever you want to pull the responses in as members.`);
}

/**
 * Builds the standalone sign-up form and records it in Document Properties.
 *
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {Form|null} The new form, or null if it could not be created.
 */
function createSignupForm(ss) {
  ss = fetchSpreadsheet(ss);
  const config = fetchProperties('configuration') || {};
  const groupName = config.groupName || `${LEAGUE} Picks Pool`;
  const year = config.year || fetchYear();

  try {
    const form = FormApp.create(`${groupName} - ${year} Sign-Up`);
    form.setDescription(`Sign up for the ${year} ${groupName}. One entry covers every pool for the season. You will get the week 1 picks form closer to kickoff -- this form is only to get you on the roster.`)
      .setAllowResponseEdits(true)
      .setLimitOneResponsePerUser(false)
      .setProgressBar(false);

    form.addTextItem()
      .setTitle(SIGNUP_Q_TEAM)
      .setHelpText(`The name that will show up on the weekly sheets and in the picks form. Pick something you will still like in December.`)
      .setRequired(true)
      .setValidation(nameValidation);

    form.addTextItem()
      .setTitle(SIGNUP_Q_MANAGER)
      .setHelpText(`Your actual name, so the commissioner knows who owns the team.`)
      .setRequired(true)
      .setValidation(nameValidation);

    form.addTextItem()
      .setTitle(SIGNUP_Q_EMAIL)
      .setHelpText(`Used only to send you each week's picks form.`)
      .setRequired(true)
      .setValidation(FormApp.createTextValidation()
        .setHelpText('Enter a valid email address.')
        .requireTextIsEmail()
        .build());

    const confirm = form.addMultipleChoiceItem();
    confirm.setTitle(SIGNUP_Q_CONFIRM)
      .setHelpText(`One entry covers the weekly pick 'ems, weekly consensus, season percentage, season points, season consensus, and the playoff pool. The commissioner will follow up about the entry fee -- do not send anything through this form.`)
      .setChoices([confirm.createChoice(`Yes, I'm in for the ${year} season`)])
      .showOtherOption(false)
      .setRequired(true);

    // Keep it beside the weekly forms rather than loose in Drive
    try {
      const folder = getFormsFolder(groupName);
      DriveApp.getFileById(form.getId()).moveTo(folder);
    } catch (err) {
      Logger.log(`⚠️ Could not move the sign-up form into the pool folder: ${err.message}`);
    }

    let publishedUrl = form.getPublishedUrl();
    try {
      publishedUrl = form.shortenFormUrl(publishedUrl);
    } catch (err) {
      Logger.log(`⚠️ Could not shorten the sign-up form URL: ${err.message}`);
    }

    saveProperties('signup', {
      formId: form.getId(),
      editUrl: form.getEditUrl(),
      publishedUrl: publishedUrl,
      created: new Date().toISOString()
    });

    Logger.log(`✅ Created sign-up form ${form.getId()}`);
    ss.toast(`Sign-up form created`,`✍️ SIGN-UP FORM READY`);
    return form;
  } catch (err) {
    Logger.log(`❌ Failed to create the sign-up form: ${err.stack}`);
    fetchUi().alert(`⚠️ SIGN-UP FORM ERROR`,`The sign-up form could not be created:\n\n${err.message}`,SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
}

/**
 * Reads the sign-up form's responses and adds them as members, after showing you
 * exactly who is about to be added. Nothing is written until you confirm.
 */
function importSignups() {
  const ui = fetchUi();
  const ss = fetchSpreadsheet();

  const signup = fetchProperties('signup');
  if (!signup || !signup.formId) {
    ui.alert(`✍️ NO SIGN-UP FORM`,`There is no sign-up form yet. Run "Create or Open Sign-Up Form" first.`,ui.ButtonSet.OK);
    return;
  }

  let form;
  try {
    form = FormApp.openById(signup.formId);
  } catch (err) {
    ui.alert(`⚠️ SIGN-UP FORM MISSING`,`The stored sign-up form could not be opened:\n\n${err.message}`,ui.ButtonSet.OK);
    return;
  }

  const config = fetchProperties('configuration') || {};
  const memberData = fetchProperties('members') || { memberOrder: [], members: {} };
  memberData.memberOrder = memberData.memberOrder || [];
  memberData.members = memberData.members || {};

  // Existing names, so re-running the import never duplicates anyone
  const taken = {};
  memberData.memberOrder.forEach(id => {
    const name = memberData.members[id] && memberData.members[id].name;
    if (name) taken[name.toString().trim().toLowerCase()] = true;
  });

  const responses = form.getResponses();
  if (!responses.length) {
    ui.alert(`✍️ NO SIGN-UPS YET`,`The sign-up form has no responses.`,ui.ButtonSet.OK);
    return;
  }

  const additions = [], duplicates = [], unusable = [];
  responses.forEach(response => {
    const answers = {};
    response.getItemResponses().forEach(item => {
      const title = item.getItem().getTitle().toString().trim().toLowerCase();
      const value = item.getResponse();
      answers[title] = (value === null || value === undefined) ? '' : value.toString().trim();
    });

    const team = answers[SIGNUP_Q_TEAM.toLowerCase()] || '';
    const manager = answers[SIGNUP_Q_MANAGER.toLowerCase()] || '';
    const email = answers[SIGNUP_Q_EMAIL.toLowerCase()] || '';

    if (!team) {
      unusable.push(manager || '(no name given)');
      return;
    }
    const key = team.toLowerCase();
    if (taken[key] || additions.some(a => a.key === key)) {
      duplicates.push(team);
      return;
    }
    additions.push({ key: key, team: team, manager: manager, email: email });
  });

  if (!additions.length) {
    ui.alert(`✍️ NOTHING NEW`,`All ${responses.length} sign-up${responses.length === 1 ? '' : 's'} already match members on your roster. Nothing to add.`,ui.ButtonSet.OK);
    return;
  }

  const preview = additions.slice(0, 25)
    .map(a => `  • ${a.team}${a.manager ? ' (' + a.manager + ')' : ''}`)
    .join('\n');
  const extra = additions.length > 25 ? `\n  ...and ${additions.length - 25} more` : '';
  const notes = [];
  if (duplicates.length) notes.push(`${duplicates.length} already on the roster, skipped`);
  if (unusable.length) notes.push(`${unusable.length} with no team name, skipped`);

  const answer = ui.alert(`📥 IMPORT ${additions.length} SIGN-UP${additions.length === 1 ? '' : 'S'}`,
    `About to add:\n\n${preview}${extra}\n\n${notes.length ? notes.join('; ') + '.\n\n' : ''}Add them to the roster?`,
    ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) {
    ss.toast(`Sign-up import canceled`,`⛔ CANCELED`);
    return;
  }

  additions.forEach(a => {
    const id = generateUniqueId();
    memberData.memberOrder.push(id);
    // Sign-ups happen before kickoff, so everyone joins at week 1 and starts unpaid
    memberData.members[id] = createNewMember(a.team, false, config, 1, { manager: a.manager, email: a.email });
  });

  saveProperties('members', memberData);

  Logger.log(`✅ Imported ${additions.length} sign-up(s): ${additions.map(a => a.team).join(', ')}`);
  ui.alert(`✅ ${additions.length} MEMBER${additions.length === 1 ? '' : 'S'} ADDED`,
    `Your roster now has ${memberData.memberOrder.length} member${memberData.memberOrder.length === 1 ? '' : 's'}.\n\nMark who has paid in the Member Manager. If you have already deployed the tracking sheets, re-run "Deploy Extra Tracking Sheets" so the new members appear on them.`,
    ui.ButtonSet.OK);
}

// ============================================================================================================================================
// CONSENSUS POOL (Pool 2 weekly, Pool 5 season)
// ============================================================================================================================================

/**
 * The group's pick for a single matchup: whichever team the majority took.
 * A dead heat returns "TIE", which the pool rules treat as +1 for the consensus
 * and +1 for every member.
 *
 * @param {Array} picks Every member's pick for one matchup.
 * @returns {string} The majority team, "TIE" on a 50/50, or "" when nobody has picked.
 */
function consensusPickForGame(picks) {
  const tally = {};
  (picks || []).forEach(pick => {
    const value = (pick === null || pick === undefined) ? '' : pick.toString().trim();
    if (value) tally[value] = (tally[value] || 0) + 1;
  });

  let leader = '', leaderCount = 0, tied = false;
  Object.keys(tally).forEach(team => {
    if (tally[team] > leaderCount) {
      leader = team;
      leaderCount = tally[team];
      tied = false;
    } else if (tally[team] === leaderCount) {
      tied = true;
    }
  });

  if (!leader) return '';
  return tied ? 'TIE' : leader;
}

/**
 * Scores the consensus row for a week against the recorded outcomes.
 * Only settled matchups count, so a part-played week reports what is known so far.
 *
 * @param {number} week The week to score.
 * @param {Spreadsheet} [ss] Optional spreadsheet handle.
 * @returns {Object|null} {correct, tieGames, settled, adjusted}, or null if the week has no sheet.
 */
function consensusWeeklyCorrect(week,ss) {
  ss = fetchSpreadsheet(ss);
  week = week || fetchWeek();

  const picksRange = ss.getRangeByName(`${LEAGUE}_CONSENSUS_PICKS_${week}`);
  const outcomeRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
  if (!picksRange || !outcomeRange) {
    Logger.log(`🤝 No consensus ranges found for week ${week}; has the weekly sheet been built?`);
    return null;
  }

  const picks = picksRange.getValues()[0];
  const outcomes = outcomeRange.getValues()[0];

  let correct = 0, tieGames = 0, settled = 0;
  for (let a = 0; a < outcomes.length; a++) {
    const outcome = (outcomes[a] === null || outcomes[a] === undefined) ? '' : outcomes[a].toString().trim();
    if (!outcome) continue;
    settled++;
    if (outcome.toUpperCase() === 'TIE') {
      tieGames++;            // the NFL tied it: freebie for the crowd, same as for every member
      continue;
    }
    const pick = (picks[a] === null || picks[a] === undefined) ? '' : picks[a].toString().trim();
    // "SPLIT" means members divided 50/50, so the crowd has no pick and scores nothing here
    if (pick && pick !== 'SPLIT' && pick === outcome) correct++;
  }

  return { correct: correct, tieGames: tieGames, settled: settled, adjusted: correct + tieGames };
}

/**
 * Whether a member beat the crowd. Matching the consensus exactly is not a win.
 * Both sides gain +1 per 50/50 matchup, so the tie adjustment never decides the
 * comparison on its own -- it is applied to keep the displayed totals honest.
 *
 * @param {number} playerCorrect The member's raw correct picks.
 * @param {number} consensusAdjusted The consensus score, tie games already included.
 * @param {number} [tieGames] Count of 50/50 matchups that week.
 * @returns {number} 1 when the member beat the consensus, otherwise 0.
 */
function playerBeatConsensus(playerCorrect, consensusAdjusted, tieGames) {
  const adjustedPlayer = Number(playerCorrect || 0) + Number(tieGames || 0);
  return adjustedPlayer > Number(consensusAdjusted || 0) ? 1 : 0;
}

// ============================================================================================================================================
// UTILITIES
// ============================================================================================================================================

/**
 * ESPN endpoints reject plain Apps Script requests from Google's servers with a 403
 * often enough to break schedule and score fetching. Every ESPN call goes through
 * these two helpers so the request carries browser-like headers and retries on the
 * transient failures (403 / 429 / 5xx) instead of dying on the first one.
 *
 * @param {string} url The ESPN endpoint to fetch.
 * @param {Object} [options] Optional UrlFetchApp params, merged over the defaults.
 * @returns {HTTPResponse} The successful response.
 * @throws {Error} If every attempt fails, with the last status code in the message.
 */
function espnFetch(url, options) {
  const params = Object.assign({
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    headers: {
      'User-Agent': ESPN_USER_AGENT,
      'Accept': 'application/json, text/plain, */*',
      'Accept-Language': 'en-US,en;q=0.9',
      'Referer': 'https://www.espn.com/',
      'Origin': 'https://www.espn.com'
    }
  }, options || {});

  let lastCode = null;
  let lastBody = '';
  for (let attempt = 1; attempt <= ESPN_FETCH_ATTEMPTS; attempt++) {
    let response;
    try {
      response = UrlFetchApp.fetch(url, params);
    } catch (err) {
      lastCode = 'exception';
      lastBody = err.message || '';
      Logger.log(`⚠️ ESPN fetch attempt ${attempt}/${ESPN_FETCH_ATTEMPTS} threw for ${url}: ${err.message}`);
      if (attempt < ESPN_FETCH_ATTEMPTS) Utilities.sleep(ESPN_FETCH_BACKOFF_MS * attempt);
      continue;
    }
    const code = response.getResponseCode();
    if (code === 200) {
      if (attempt > 1) Logger.log(`✅ ESPN fetch recovered on attempt ${attempt} for ${url}`);
      return response;
    }
    lastCode = code;
    lastBody = response.getContentText().slice(0, 200);
    // 4xx other than 403/429 will not fix themselves with a retry
    if (code !== 403 && code !== 429 && code < 500) {
      break;
    }
    Logger.log(`⚠️ ESPN fetch attempt ${attempt}/${ESPN_FETCH_ATTEMPTS} got ${code} for ${url}`);
    if (attempt < ESPN_FETCH_ATTEMPTS) Utilities.sleep(ESPN_FETCH_BACKOFF_MS * attempt);
  }
  Logger.log(`❌ ESPN fetch failed for ${url} (last status: ${lastCode}) ${lastBody}`);
  throw new Error(`ESPN request failed for ${url} (status ${lastCode})`);
}

/**
 * Fetches an ESPN endpoint and returns the parsed JSON body.
 * Throws on transport failure or unparseable body, matching the behavior callers
 * already handle (a bare JSON.parse of an error page threw too).
 *
 * @param {string} url The ESPN endpoint to fetch.
 * @param {Object} [options] Optional UrlFetchApp params.
 * @returns {Object} The parsed JSON payload.
 */
function espnFetchJson(url, options) {
  const text = espnFetch(url, options).getContentText();
  try {
    return JSON.parse(text);
  } catch (err) {
    Logger.log(`❌ ESPN returned an unparseable body for ${url}: ${text.slice(0, 200)}`);
    throw new Error(`ESPN returned a non-JSON response for ${url}`);
  }
}

/**
 * Picks the game a week's tiebreaker questions should be based on.
 *
 * Pool rules: use the late Monday night game -- the later one when a week has two
 * Monday games. Playoff weeks have no Monday game, so this falls back to the last
 * kickoff of the week, which is what the playoff rules ask for anyway.
 *
 * DAY[n].index is already chronological across a football week
 * (Wed -4 -> Thu -3 -> Fri -2 -> Sat -1 -> Sun 0 -> Mon 1), so sorting on it
 * handles the Wednesday/Thursday openers without a special case.
 *
 * @param {Array<Object>} games The week's games, as stored in gamePlan.games.
 * @returns {Object|null} The chosen game, or null when there are no games.
 */
function getTiebreakerGame(games) {
  if (!games || games.length === 0) return null;

  const kickoff = (game) => {
    let dayIndex = 0;
    if (game.day !== undefined && DAY[game.day]) {
      dayIndex = DAY[game.day].index;
    } else if (game.dayName) {
      const match = Object.keys(DAY).find(key => DAY[key].name === game.dayName);
      if (match !== undefined) dayIndex = DAY[match].index;
    }
    return (dayIndex * 1440) + ((Number(game.hour) || 0) * 60) + (Number(game.minute) || 0);
  };

  const mondayGames = games.filter(game => game.dayName === 'Monday');
  const pool = mondayGames.length > 0 ? mondayGames : games;

  let latest = pool[0];
  for (const game of pool) {
    // >= so a later entry in the schedule wins an exact time tie
    if (kickoff(game) >= kickoff(latest)) latest = game;
  }
  Logger.log(`⚖️ Tiebreaker game selected: ${latest.awayTeam} @ ${latest.homeTeam} (${latest.dayName}, ${formatTime(latest.hour, latest.minute)})`);
  return latest;
}

/**
 * Displays a clean modal dialog with a link for the user to click.
 * This is the standard way to direct a user to a URL from a server-side script.
 *
 * @param {string} url The URL the link should point to.
 * @param {string} title The title for the dialog window.
 * @param {string} linkText The text to display for the link itself.
 */
function showLinkDialog(url, title, linkText, subText) {
  const htmlContent = `
    <div style="font-family: 'Montserrat', sans-serif; text-align: center; padding: 20px;">
      <p style="font-size: 16px;">
        <a href="${url}" target="_blank" onclick="google.script.host.close()" 
           style="font-weight: bold; text-decoration: none; background-color: #013369; color: white; padding: 10px 20px; border-radius: 5px;">
          ${linkText}
        </a>
      </p>
      ${subText ? '<div style="font-size: 12px;">' + subText + '</div>' : ''}
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/**
 * Displays a polished dialog with "Edit," "Open," and "Copy Link"
 * action buttons for a newly created form.
 *
 * @param {Object} newFormsData An object containing the form's URLs and other details.
 * @param {number} week The week number for which the form was created.
 */
function showFormActionsDialog(newFormsData, week) {
  // Destructure the URLs from the input object for easy access
  const { editUrl, publishedUrl } = newFormsData;
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Montserrat', Arial, sans-serif; padding: 10px; text-align: center; }
        h2 { color: #013369; margin-top: 0; }
        p { font-size: 14px; color: #333; }
        .button-row { display: flex; gap: 10px; justify-content: center; margin-top: 20px; }
        .btn { padding: 8px 15px;  font-size: 14px;  font-weight: 600;  border: none;  border-radius: 5px;  cursor: pointer;  color: white;  text-decoration: none; display: flex;  align-items: center;  justify-content: center;  gap: 5px; }
        .btn-primary { background-color: #013369; }
        .btn-primary:hover { background-color: #2067b3; }
        .btn-secondary { background-color: #878787; }
        .btn-secondary:hover { background-color: #A8A8A8; }
        .btn-alert { background-color: #ff913d; }
        .btn-alert:hover { background-color: #e86705; }
        button.btn-alert { color: white; }
      </style>
    </head>
    <body>
      <h2>✅ Success!</h2>
      <p>Your form for Week ${week} has been created. Your form ${newFormsData.gamePlan.membershipLocked ? 'is open to new members.' : 'only accepts submissions from existing members.'} Copy and share the link with your pool members.</p>
      <div class="button-row">
        <a href="${editUrl}" target="_blank" class="btn btn-secondary">📝 Edit Form</a>
        <a href="${publishedUrl}" target="_blank" class="btn btn-primary">📂 Open Form</a>
        <button id="copy-btn" class="btn btn-alert">📤 Copy Link</button>
      </div>
      <script>
        document.getElementById('copy-btn').onclick = function() {
          const linkToCopy = "${publishedUrl}";
          const button = this;
          navigator.clipboard.writeText(linkToCopy).then(() => {
            button.textContent = '✅ Copied!';
            button.disabled = true;
            setTimeout(() => {
              button.textContent = '📤 Copy Link';
              button.disabled = false;
            }, 2000);
          }).catch(err => {
            Logger.log('Failed to copy: ', err);
            prompt("Please copy this link manually:", linkToCopy);
          });
        };
      </script>
    </body>
    </html>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(450).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Week ${week} Form Created`);
}


/**
 * Generates a simple, robust, and sufficiently unique ID string.
 * Creates an ID like "id_1234567890".
 *
 * @returns {string} A new unique ID.
 */
function generateUniqueId() {
  const randomPart = Math.random().toString(36).substring(3, 15).toUpperCase();
  return `id_${randomPart}`;
}

// RESET Function to reset and create menu for runFirst
function resetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(`↩️ Return to spreadsheet for prompts`);
  let prompt = ui.alert(`❗ RESET?`,'Reset spreadsheet and delete all data?', ui.ButtonSet.YES_NO);
  if (prompt == 'YES') {
    
    let promptTwo = ui.alert('Are you sure? This would be very difficult to recover from.',ui.ButtonSet.YES_NO);
    if (promptTwo == 'YES') {
      let ranges = ss.getNamedRanges();
      for (let a = 0; a < ranges.length; a++){
        ranges[a].remove();
      }
      let sheets = ss.getSheets();
      let baseSheet = ss.insertSheet();
      for (let a = 0; a < sheets.length; a++){
        ss.deleteSheet(sheets[a]);
      }
      let protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }
      protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }      
      baseSheet.setName('Sheet1');

      // Deletes initialization, time zone, and any other response-associated properites
      let properties = PropertiesService.getDocumentProperties();
      properties.deleteAllProperties();

      deleteTriggers();

      initializeMenu();

    } else {
      ss.toast('Canceled reset');
    }
  } else {
    ss.toast('Canceled reset');
  }
  
}

// FETCH SPREADSHEET - Checks that the 'ss' variable passed into a script is not null, undefined, or a non-spreadsheet
function fetchSpreadsheet(ss) {
  try {
    if (!ss) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
    if (ss && typeof ss.getSheets === 'function' && typeof ss.getId === 'function') {
      return ss;
    } else {
      throw new Error('Invalid Spreadsheet object');
    }
  } catch (err) {
    if (ss !== null && ss !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ss + ' \'' + ss + '\' to the \'fetchSpreadsheet\' function.');
      Logger.log(err.stack);
    }
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return ss;
}

// FETCH UI - Checks that the 'ui' variable passed into a script is not null, undefined, or a non-UI
function fetchUi(ui) {
  try{
    if (typeof ui.showModalDialog !== 'function') {
      throw new Error('Non-UI passed');
    }
  }
  catch (err) {
    if (ui !== null && ui !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ui + ' \'' + ui + '\' to the \'fetchUi\' function.');
    }
    ui = SpreadsheetApp.getUi();
  }
  return ui;
}

// SERVICE Function to remove all triggers on project
function deleteTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let a = 0; a < triggers.length; a++) {
    ScriptApp.deleteTrigger(triggers[a]);
  }
}

// ADJUST ROWS - Cleans up rows of a sheet by providing the total rows that currently exist with data
function adjustRows(sheet,rows,verbose){
  let maxRows = sheet.getMaxRows(); 
  if (rows == undefined || rows == null) {
    rows = sheet.getLastRow();
  }
  if (rows > 0 && rows > maxRows) {
    sheet.insertRowsAfter(maxRows,(rows-maxRows));
    if(verbose) return Logger.log('Added ' + (rows-maxRows) + ' rows');
  } else if (rows < maxRows && rows != 0){
    sheet.deleteRows((rows+1), (maxRows-rows));
    if(verbose) return Logger.log('Removed ' + (maxRows - rows) + ' rows');
  } else {
    if(verbose) return Logger.log('Rows not adjusted');
  }
}

// ADJUST COLUMNS - Cleans up columns of a sheet by providing the total columns that currently exist with data
function adjustColumns(sheet,columns,verbose){
  let maxColumns = sheet.getMaxColumns(); 
  if (columns == undefined || columns == null) {
    columns = sheet.getLastColumn();
  }
  if (columns > 0 && columns > maxColumns) {
    sheet.insertColumnsAfter(maxColumns,(columns-maxColumns));
    if(verbose) return Logger.log('Added ' + (columns-maxColumns) + ' columns');
  }  else if (columns < maxColumns && columns != 0){
    sheet.deleteColumns((columns+1), (maxColumns-columns));
    if(verbose) return Logger.log('Removed ' + (maxColumns - columns) + ' column(s)');
  } else {
    if(verbose) return Logger.log('Columns not adjusted');
  }
}

// GENERATES HEX GRADIENT - Provide a start and end and a count of values and this function generates a HEX gradient. Midpoint value is optional.
function hexGradient(start, end, count, midpoint) { // start and end in either 3 or 6 digit hex values, count is total values in array to return
  if (count < 2 || count.isNaN) {
    Logger.log('ERROR: Please provide a \'count\' value of 2 or greater');
    return null;
  } else {
    count = Math.ceil(count);
    if (midpoint == null || midpoint == undefined) {
      // strip the leading # if it's there
      start = start.replace(/^\s*#|\s*$/g, '');
      end = end.replace(/^\s*#|\s*$/g, '');

      // convert 3 char codes --> 6, e.g. `E0F` --> `EE00FF`
      if(start.length == 3){
        start = start.replace(/(.)/g, '$1$1');
      }

      if(end.length == 3){
        end = end.replace(/(.)/g, '$1$1');
      }

      let arr = ['#'+start];
      let tmpRed, tmpGreen, tmpBlue;

      // get colors
      let startRed = parseInt(start.substr(0, 2), 16),
          startGreen = parseInt(start.substr(2, 2), 16),
          startBlue = parseInt(start.substr(4, 2), 16);
      let endRed = parseInt(end.substr(0, 2), 16),
          endGreen = parseInt(end.substr(2, 2), 16),
          endBlue = parseInt(end.substr(4, 2), 16);
      let stepRed = (endRed-startRed)/(count-1),
          stepGreen = (endGreen-startGreen)/(count-1),
          stepBlue = (endBlue-startBlue)/(count-1);
      
      for (let a = 1; a < count-1; a++) {
        // calculate the step differential for each color
        tmpRed = ((stepRed * a) + startRed).toString(16).split('.')[0];
        tmpGreen = ((stepGreen * a) + startGreen).toString(16).split('.')[0];
        tmpBlue = ((stepBlue * a) + startBlue).toString(16).split('.')[0];
        // ensure 2 digits by color
        if( tmpRed.length == 1 ) tmpRed = '0' + tmpRed;
        if( tmpGreen.length == 1 ) tmpGreen = '0' + tmpGreen;
        if( tmpBlue.length == 1 ) tmpBlue = '0' + tmpBlue;
        arr.push(('#' + tmpRed + tmpGreen + tmpBlue).toUpperCase());
      }
      arr.push('#'+end);
      return arr;
    } else {
      count = Math.ceil(count);
      if (count % 2 == 0) {
        count++;
        // Logger.log('Even number provided with midpoint, increasing count to ' + count);
      }
      let half = Math.ceil(count/2);
      let arr = hexGradient(start,midpoint,half);
      arr.pop();
      let arr2 = hexGradient(midpoint,end,half);
      arr = arr.concat(arr2);
      return arr;
    }
  }
}

// ENSURE ARRAY IS RECTANGULAR - a function to ensure that an array has blank values if it fails to have a full set of columns per row
function makeArrayRectangular(arr) {
  const maxLength = Math.max(...arr.map(row => row.length));
  for (let a = 0; a < arr.length; a++) {
    // While the row's length is less than the maximum length, push a placeholder value
    while (arr[a].length < maxLength) {
      arr[a].push('');
    }
  }
  return arr;
}

// GET TIMEZONE
function timezoneSet() {
  // Get the value for the script property timezone
  const scriptProperties = PropertiesService.getDocumentProperties();
  const tz = scriptProperties.getProperty('tz');
  if (tz != null) {
    return true;
  } else {
    Logger.log('No timezone confirmation has been done yet');
    return false;
  }
}

// SET PROPRTY - sets a script property based on an inputted name (string) and a value (string/array/object) (essentially this ia global variable)
function setProperty(property,value){
  const scriptProperties = PropertiesService.getDocumentProperties();
  if (typeof value === 'object' && !Array.isArray(value) && value !== null) {
    scriptProperties.setProperty(property,JSON.stringify(value));
  } else {
    scriptProperties.setProperty(property,value);
  }
}

// OPEN URL - Quick script to open a new tab with the newly created form, in this case
function openUrl(url,week){
  if (!url || typeof url !== 'string') {
    throw new Error("Invalid URL provided.");
  }
  if (week == null) {
    week = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WEEK').getValue();
  }
  if (week == undefined) {
    week = fetchWeek();
  }

  // Create the HTML content with the Montserrat font
  let htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap" rel="stylesheet">
    <div style="font-family: 'Montserrat', sans-serif; text-align: center; padding: 20px;">
      <p style="font-size: 22px;"><a href="${url}" target="_blank" style="font-weight: bold;">Click for Week ` + week + ` Form</a></p>
    </div>
  `;

  let htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(350)
    .setHeight(180);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

/**
 * A simple toast message helper.
 */
function showToast(message,ss) {
  ss = fetchSpreadsheet(ss);
  ss.toast(message);
}


// ============================================================================================================================================
// DEBUG TOOLS
// ============================================================================================================================================


// FRONTEND CONFIGURATION REVIEW - Brings up an alert in the Google Sheet to show what configuration variables are set
// This is a back-end and unused script for general usage
function checkDocumentConfiguration() {
  Logger.log(`Fetching configuration. Please review report in Sheets tab...`);
  try {
    const ss = fetchSpreadsheet();
    const ui = fetchUi();
    let str = '';
    const props = fetchProperties('configuration');
    Object.keys(props).forEach(key => {
      let subStr = '\n';
      if (typeof props[key] == 'object') {
        subStr += key + ': {\n';
        Object.keys(props[key]).forEach(subKey => {
          if (typeof props[key][subKey] == 'object') {
            subStr += '-' + subKey + ': {\n';
            Object.keys(props[key][subKey]).forEach(subSubKey => {
              subStr += '--' + subSubKey + ': ' + props[key][subKey][subSubKey] + '\n';
            });
            subStr += '--}\n';
          } else {
            subStr += '-' + subKey + ': ' + props[key][subKey] + '\n';
          }
        });
        subStr += '-}\n';
      } else {
          subStr += '-' + key + ': ' + props[key] + '\n';
      }
      str += (key + ': ' + (typeof props[key] === 'boolean' ? (props[key] ? '✅\n' : '❌\n') : (typeof props[key] == 'object' ? subStr : props[key] + '\n' )));
    });

    ui.alert(str,ui.ButtonSet.OK);
  } catch (err) {
    Logger.log(`❌ Failed to retrieve members sidebar data: ${err.stack}`);
    return { properties: {} };
  }
}

// VIEW USER PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewUserProperties() {
  let userProperties = PropertiesService.getUserProperties().getProperties();
  Logger.log(`User Properties:`);
  for (let key in userProperties) {
    Logger.log(key + ': ' + userProperties[key]);
  }
}

// VIEW SCRIPT PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewScriptProperties() {
  let scriptProperties = PropertiesService.getScriptProperties().getProperties();
  Logger.log('Script Properties:');
  for (let key in scriptProperties) {
    Logger.log(key + ': ' + scriptProperties[key]);
  }
}

// VIEW DOCUMENT PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewDocumentProperties() {
  let documentProperties = PropertiesService.getDocumentProperties().getProperties();
  Logger.log('Document Properties:');
  for (let key in documentProperties) {
    Logger.log(key + ': ' + documentProperties[key]);
  }
}

// VIEW RESPONSES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewResponseJSON(week) {
  week = week || 1
  let docProps = PropertiesService.getDocumentProperties();
  let formsData = JSON.parse(docProps.getProperty('forms')) || {};
  const databaseSheet = getDatabaseSheet();
  const responseSheet = databaseSheet.getSheetByName(`WK${week}`);
  
  // Parse the latest, de-duplicated picks from the response sheet
  const memberData = JSON.parse(docProps.getProperty('members')) || {};
  const parsedPicks = parseAllPicksFromSheet(responseSheet, memberData);
  Logger.log(`Week ${week} Pick Responses:`);
  for (let key in parsedPicks) {
    Logger.log(key + ': ' + JSON.stringify(parsedPicks[key]));
  }
}

/**
 * Opens the Manual Member Data Editor modal.
 */
function memberDataPanel() {
  const html = HtmlService.createHtmlOutputFromFile('memberDataPanel')
      .setWidth(1300)
      .setHeight(850)
      .setTitle('Master JSON Member Editor');
  SpreadsheetApp.getUi().showModalDialog(html, 'Master JSON Member Editor');
}

function getManualEditorData() {
  return {
    memberData: JSON.parse(PropertiesService.getDocumentProperties().getProperty('members')) || {},
    config: JSON.parse(PropertiesService.getDocumentProperties().getProperty('configuration')) || {}
  };
}

function saveManualMemberData(payload) {
  try {
    const memberData = JSON.parse(PropertiesService.getDocumentProperties().getProperty('members')) || {};

    const toArray = (str) => {
      if (!str) return [];
      return str.split(',').map(v => {
        v = v.trim();
        if (v === 'null' || v === '') return null;
        return isNaN(v) ? v : parseInt(v);
      });
    };

    for (const id in payload) {
      const p = payload[id];
      const m = memberData.members[id];
      if (!m) continue;

      m.name = p.name;
      m.active = (p.active === true);
      m.paid = (p.paid === true);

      // Process Survivor
      if (p.sP !== undefined) {
        m.sP = toArray(p.sP);
        m.sL = toArray(p.sL);
        m.sR = toArray(p.sR);
        m.sE = p.sE; // This is now a clean array of booleans from the client
        m.sO = p.sO === "" ? null : parseInt(p.sO);
      }

      // Process Eliminator
      if (p.eP !== undefined) {
        m.eP = toArray(p.eP);
        m.eL = toArray(p.eL);
        m.eR = toArray(p.eR);
        m.eE = p.eE; // Array of booleans
        m.eO = p.eO === "" ? null : parseInt(p.eO);
      }
    }

    saveProperties('members', memberData);
    return "✅ Master JSON data saved successfully!";
  } catch (err) {
    throw new Error(err.message);
  }
}
