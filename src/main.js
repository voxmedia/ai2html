// ================================================
// Constant data
// ================================================

// html entity substitution
// not used ==> ["\x22","&quot;"], ["\x3C","&lt;"], ["\x3E","&gt;"], ["\x26","&amp;"],
var htmlCharacterCodes = [["\xA0","&nbsp;"], ["\xA1","&iexcl;"], ["\xA2","&cent;"], ["\xA3","&pound;"], ["\xA4","&curren;"], ["\xA5","&yen;"], ["\xA6","&brvbar;"], ["\xA7","&sect;"], ["\xA8","&uml;"], ["\xA9","&copy;"], ["\xAA","&ordf;"], ["\xAB","&laquo;"], ["\xAC","&not;"], ["\xAD","&shy;"], ["\xAE","&reg;"], ["\xAF","&macr;"], ["\xB0","&deg;"], ["\xB1","&plusmn;"], ["\xB2","&sup2;"], ["\xB3","&sup3;"], ["\xB4","&acute;"], ["\xB5","&micro;"], ["\xB6","&para;"], ["\xB7","&middot;"], ["\xB8","&cedil;"], ["\xB9","&sup1;"], ["\xBA","&ordm;"], ["\xBB","&raquo;"], ["\xBC","&frac14;"], ["\xBD","&frac12;"], ["\xBE","&frac34;"], ["\xBF","&iquest;"], ["\xD7","&times;"], ["\xF7","&divide;"], ["\u0192","&fnof;"], ["\u02C6","&circ;"], ["\u02DC","&tilde;"], ["\u2002","&ensp;"], ["\u2003","&emsp;"], ["\u2009","&thinsp;"], ["\u200C","&zwnj;"], ["\u200D","&zwj;"], ["\u200E","&lrm;"], ["\u200F","&rlm;"], ["\u2013","&ndash;"], ["\u2014","&mdash;"], ["\u2018","&lsquo;"], ["\u2019","&rsquo;"], ["\u201A","&sbquo;"], ["\u201C","&ldquo;"], ["\u201D","&rdquo;"], ["\u201E","&bdquo;"], ["\u2020","&dagger;"], ["\u2021","&Dagger;"], ["\u2022","&bull;"], ["\u2026","&hellip;"], ["\u2030","&permil;"], ["\u2032","&prime;"], ["\u2033","&Prime;"], ["\u2039","&lsaquo;"], ["\u203A","&rsaquo;"], ["\u203E","&oline;"], ["\u2044","&frasl;"], ["\u20AC","&euro;"], ["\u2111","&image;"], ["\u2113",""], ["\u2116",""], ["\u2118","&weierp;"], ["\u211C","&real;"], ["\u2122","&trade;"], ["\u2135","&alefsym;"], ["\u2190","&larr;"], ["\u2191","&uarr;"], ["\u2192","&rarr;"], ["\u2193","&darr;"], ["\u2194","&harr;"], ["\u21B5","&crarr;"], ["\u21D0","&lArr;"], ["\u21D1","&uArr;"], ["\u21D2","&rArr;"], ["\u21D3","&dArr;"], ["\u21D4","&hArr;"], ["\u2200","&forall;"], ["\u2202","&part;"], ["\u2203","&exist;"], ["\u2205","&empty;"], ["\u2207","&nabla;"], ["\u2208","&isin;"], ["\u2209","&notin;"], ["\u220B","&ni;"], ["\u220F","&prod;"], ["\u2211","&sum;"], ["\u2212","&minus;"], ["\u2217","&lowast;"], ["\u221A","&radic;"], ["\u221D","&prop;"], ["\u221E","&infin;"], ["\u2220","&ang;"], ["\u2227","&and;"], ["\u2228","&or;"], ["\u2229","&cap;"], ["\u222A","&cup;"], ["\u222B","&int;"], ["\u2234","&there4;"], ["\u223C","&sim;"], ["\u2245","&cong;"], ["\u2248","&asymp;"], ["\u2260","&ne;"], ["\u2261","&equiv;"], ["\u2264","&le;"], ["\u2265","&ge;"], ["\u2282","&sub;"], ["\u2283","&sup;"], ["\u2284","&nsub;"], ["\u2286","&sube;"], ["\u2287","&supe;"], ["\u2295","&oplus;"], ["\u2297","&otimes;"], ["\u22A5","&perp;"], ["\u22C5","&sdot;"], ["\u2308","&lceil;"], ["\u2309","&rceil;"], ["\u230A","&lfloor;"], ["\u230B","&rfloor;"], ["\u2329","&lang;"], ["\u232A","&rang;"], ["\u25CA","&loz;"], ["\u2660","&spades;"], ["\u2663","&clubs;"], ["\u2665","&hearts;"], ["\u2666","&diams;"]];

// Add to the fonts array to make the script work with your own custom fonts.
// To make it easier to add to this array, use the "fonts" worksheet of this Google Spreadsheet:
  // https://docs.google.com/spreadsheets/d/13ESQ9ktfkdzFq78FkWLGaZr2s3lNbv2cN25F2pYf5XM/edit?usp=sharing
  // Make a copy of the spreadsheet for yourself.
  // Modify the settings to taste.
var fonts = [
  {"aifont":"ArialMT","family":"arial,helvetica,sans-serif","weight":"","style":""},
  {"aifont":"Arial-BoldMT","family":"arial,helvetica,sans-serif","weight":"bold","style":""},
  {"aifont":"Arial-ItalicMT","family":"arial,helvetica,sans-serif","weight":"","style":"italic"},
  {"aifont":"Arial-BoldItalicMT","family":"arial,helvetica,sans-serif","weight":"bold","style":"italic"},
  {"aifont":"Georgia","family":"georgia,'times new roman',times,serif","weight":"","style":""},
  {"aifont":"Georgia-Bold","family":"georgia,'times new roman',times,serif","weight":"bold","style":""},
  {"aifont":"Georgia-Italic","family":"georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"Georgia-BoldItalic","family":"georgia,'times new roman',times,serif","weight":"bold","style":"italic"},
  {"aifont":"NYTFranklin-Light","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"NYTFranklin-Medium","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"500","style":""},
  {"aifont":"NYTFranklin-SemiBold","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"600","style":""},
  {"aifont":"NYTFranklin-Semibold","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"600","style":""},
  {"aifont":"NYTFranklinSemiBold-Regular","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"600","style":""},
  {"aifont":"NYTFranklin-SemiboldItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"600","style":"italic"},
  {"aifont":"NYTFranklin-Bold","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"NYTFranklin-LightItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"300","style":"italic"},
  {"aifont":"NYTFranklin-MediumItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"500","style":"italic"},
  {"aifont":"NYTFranklin-BoldItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"700","style":"italic"},
  {"aifont":"NYTFranklin-ExtraBold","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"800","style":""},
  {"aifont":"NYTFranklin-ExtraBoldItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"800","style":"italic"},
  {"aifont":"NYTFranklin-Headline","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"bold","style":""},
  {"aifont":"NYTFranklin-HeadlineItalic","family":"nyt-franklin,arial,helvetica,sans-serif","weight":"bold","style":"italic"},
  {"aifont":"NYTCheltenham-ExtraLight","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"200","style":""},
  {"aifont":"NYTCheltenhamExtLt-Regular","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"200","style":""},
  {"aifont":"NYTCheltenham-Light","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"300","style":""},
  {"aifont":"NYTCheltenhamLt-Regular","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"300","style":""},
  {"aifont":"NYTCheltenham-LightSC","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"300","style":""},
  {"aifont":"NYTCheltenham-Book","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"400","style":""},
  {"aifont":"NYTCheltenhamBook-Regular","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"400","style":""},
  {"aifont":"NYTCheltenham-Wide","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":""},
  {"aifont":"NYTCheltenhamMedium-Regular","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"500","style":""},
  {"aifont":"NYTCheltenham-Medium","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"500","style":""},
  {"aifont":"NYTCheltenham-Bold","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"700","style":""},
  {"aifont":"NYTCheltenham-BoldCond","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"bold","style":""},
  {"aifont":"NYTCheltenhamCond-BoldXC","family":"nyt-cheltenham-extra-cn-bd,georgia,'times new roman',times,serif","weight":"bold","style":""},
  {"aifont":"NYTCheltenham-BoldExtraCond","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"bold","style":""},
  {"aifont":"NYTCheltenham-ExtraBold","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"bold","style":""},
  {"aifont":"NYTCheltenham-ExtraLightIt","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-ExtraLightItal","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-LightItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-BookItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-WideItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-MediumItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"","style":"italic"},
  {"aifont":"NYTCheltenham-BoldItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"700","style":"italic"},
  {"aifont":"NYTCheltenham-ExtraBoldItal","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"bold","style":"italic"},
  {"aifont":"NYTCheltenham-ExtraBoldItalic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"bold","style":"italic"},
  {"aifont":"NYTCheltenhamSH-Regular","family":"nyt-cheltenham-sh,nyt-cheltenham,georgia,'times new roman',times,serif","weight":"400","style":""},
  {"aifont":"NYTCheltenhamSH-Italic","family":"nyt-cheltenham-sh,nyt-cheltenham,georgia,'times new roman',times,serif","weight":"400","style":"italic"},
  {"aifont":"NYTCheltenhamSH-Bold","family":"nyt-cheltenham-sh,nyt-cheltenham,georgia,'times new roman',times,serif","weight":"700","style":""},
  {"aifont":"NYTCheltenhamSH-BoldItalic","family":"nyt-cheltenham-sh,nyt-cheltenham,georgia,'times new roman',times,serif","weight":"700","style":"italic"},
  {"aifont":"NYTCheltenhamWide-Regular","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"500","style":""},
  {"aifont":"NYTCheltenhamWide-Italic","family":"nyt-cheltenham,georgia,'times new roman',times,serif","weight":"500","style":"italic"},
  {"aifont":"NYTKarnakText-Regular","family":"nyt-karnak-display-130124,georgia,'times new roman',times,serif","weight":"400","style":""},
  {"aifont":"NYTKarnakDisplay-Regular","family":"nyt-karnak-display-130124,georgia,'times new roman',times,serif","weight":"400","style":""},
  {"aifont":"NYTStymieLight-Regular","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"NYTStymieMedium-Regular","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"500","style":""},
  {"aifont":"StymieNYT-Light","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"StymieNYT-LightPhoenetic","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"StymieNYT-Lightitalic","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":"italic"},
  {"aifont":"StymieNYT-Medium","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"500","style":""},
  {"aifont":"StymieNYT-MediumItalic","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"500","style":"italic"},
  {"aifont":"StymieNYT-Bold","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"StymieNYT-BoldItalic","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":"italic"},
  {"aifont":"StymieNYT-ExtraBold","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"StymieNYT-ExtraBoldText","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"StymieNYT-ExtraBoldTextItal","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":"italic"},
  {"aifont":"StymieNYTBlack-Regular","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"StymieBT-ExtraBold","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"700","style":""},
  {"aifont":"Stymie-Thin","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"Stymie-UltraLight","family":"nyt-stymie,arial,helvetica,sans-serif","weight":"300","style":""},
  {"aifont":"NYTMagSans-Regular","family":"'nyt-mag-sans',arial,helvetica,sans-serif","weight":"500","style":""},
  {"aifont":"NYTMagSans-Bold","family":"'nyt-mag-sans',arial,helvetica,sans-serif","weight":"700","style":""}
];

// CSS text-transform equivalents
var caps = [
  {"ai":"FontCapsOption.NORMALCAPS","html":"none"},
  {"ai":"FontCapsOption.ALLCAPS","html":"uppercase"},
  {"ai":"FontCapsOption.SMALLCAPS","html":"uppercase"}
];

// CSS text-align equivalents
var align = [
  {"ai":"Justification.LEFT","html":"left"},
  {"ai":"Justification.RIGHT","html":"right"},
  {"ai":"Justification.CENTER","html":"center"},
  {"ai":"Justification.FULLJUSTIFY","html":"justify"},
  {"ai":"Justification.FULLJUSTIFYLASTLINELEFT","html":"justify"},
  {"ai":"Justification.FULLJUSTIFYLASTLINECENTER","html":"justify"},
  {"ai":"Justification.FULLJUSTIFYLASTLINERIGHT","html":"justify"}
];

var blendModes = [
  {ai: "BlendModes.MULTIPLY", html: "multiply"}
];

// list of CSS properties used for translating AI text styles
var cssTextStyleProperties = [
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'color',
  'line-height',
  'letter-spacing',
  'opacity',
  'padding-top',
  'padding-bottom',
  'text-align',
  'text-transform',
  'mix-blend-mode'
];

var nyt5Breakpoints = [
  { name:"xsmall"    , lowerLimit:   0, upperLimit: 180 },
  { name:"small"     , lowerLimit: 180, upperLimit: 300 },
  { name:"smallplus" , lowerLimit: 300, upperLimit: 460 },
  { name:"submedium" , lowerLimit: 460, upperLimit: 600 },
  { name:"medium"    , lowerLimit: 600, upperLimit: 720 },
  { name:"large"     , lowerLimit: 720, upperLimit: 945 },
  { name:"xlarge"    , lowerLimit: 945, upperLimit:1050 },
  { name:"xxlarge"   , lowerLimit:1050, upperLimit:1600 }
];

// TODO: add these to settings spreadsheet
var nameSpace           = "g-";
var cssPrecision        = 4;
// If all three RBG channels (0-255) are below this value, convert text fill to pure black.
var rgbBlackThreshold  = 36;
var showDebugMessages  = true; // add text logged with message() function to completion alert

// ================================
// Variable declarations
// ================================

// vars to hold warnings and informational messages at the end
var feedback = [];
var warnings = [];
var errors   = [];
var oneTimeWarnings = [];
var textFramesToUnhide = [];
var objectsToRelock = [];

// Global variables set by main()
var docSettings = {};
var ai2htmlBaseSettings;
var previewProjectType, scriptEnvironment;
var doc, docPath, docName, docIsSaved;
var pBar, T;

// ===========================================================
// If running in Node.js: export functions for testing and exit
// ===========================================================
if (runningInNode()) {
  [ testBoundsIntersection,
    trim,
    stringToLines,
    contains,
    arraySubtract,
    firstBy,
    zeroPad,
    roundTo,
    folderExists,
    formatCss,
    getCssColor,
    readGitConfigFile,
    readYamlConfigFile,
    applyTemplate,
    cleanText,
    findHtmlTag,
    cleanHtmlTags,
    convertSettingsToYaml,
    parseDataAttributes,
    parseArtboardName,
    initDocumentSettings
  ].forEach(function(f) {
    module.exports[f.name] = f;
  });
  return;
}

if (!isTestedIllustratorVersion(app.version)) {
  warnings.push("Ai2html has not been tested on this version of Illustrator.");
}

if (!app.documents.length) {
  errors.push("No documents are open");

} else if (!String(app.activeDocument.path)) {
  errors.push("You need to save your Illustrator file before running this script");

} else if (app.activeDocument.documentColorSpace != DocumentColorSpace.RGB) {
  errors.push("You should change the document color mode to \"RGB\" before running ai2html (File>Document Color Mode>RGB Color).");

} else if (app.activeDocument.activeLayer.name == "Isolation Mode") {
  errors.push("Ai2html is unable to run because the document is in Isolation Mode.");

} else if (app.activeDocument.activeLayer.name == "<Opacity Mask>" && app.activeDocument.layers.length == 1) {
  // TODO: find a better way to detect this condition (mask can be renamed)
  errors.push("Ai2html is unable to run because you are editing an Opacity Mask.");

} else {
  // initialize global variables
  doc = app.activeDocument;
  docPath = doc.path + "/";
  docIsSaved = doc.saved;
  scriptEnvironment = detectScriptEnvironment();
  initDocumentSettings(scriptEnvironment);
}

if (errors.length === 0) {
  validateArtboardNames();
  initUtilityFunctions(); // init T and JSON
  T.start();
  try {
    render();
  } catch(e) {
    errors.push(formatError(e));
  }
  restoreDocumentState();
}


// ==========================================
// Save the AI document (if needed)
// ==========================================

if (docIsSaved) {
  // If document was originally in a saved state, reset the document's
  // saved flag (the document goes to unsaved state during the script,
  // because of unlocking / relocking of objects
  doc.saved = true;
} else if (errors.length === 0) {
  // Auto-save the document if no errors occurred
  var saveOptions = new IllustratorSaveOptions();
  saveOptions.pdfCompatible = false;
  doc.saveAs(new File(docPath + doc.name), saveOptions);
  feedback.push("Your Illustrator file was saved.");
}

if (pBar) pBar.close();
if (T) message("Script ran in", (T.stop() / 1000).toFixed(1), "seconds");

// =========================================================
// Show alert box, optionally prompt to generate promo image
// =========================================================

if (isTrue(docSettings.show_completion_dialog_box ) || errors.length > 0) {
  var promptForPromo = errors.length === 0 && isTrue(docSettings.write_image_files) &&
    (previewProjectType == "ai2html" || isTrue(docSettings.create_promo_image));
  var showPromo = showCompletionAlert(promptForPromo);
  if (showPromo) createPromoImage(docSettings);
}
