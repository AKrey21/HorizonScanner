/*********************************
 * Config.gs — shared constants
 *********************************/

// Sheet names
const CONTROL_SHEET = "ThemeRules";
const FEEDS_SHEET   = "RSS Feeds";
const RAW_SHEET     = "Raw Articles";
const PICKS_SHEET   = "Weekly Picks";

// Ingest behavior
const INGEST_ALL_ARTICLES = false;   // set true to ingest even if no rule matches

const INGEST_PRUNE_DAYS = 14;      // daily raw prune retention window
