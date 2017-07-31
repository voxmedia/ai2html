// =================================
// ai2html render function
// =================================

function render() {
  // Fix for issue #50
  // If a text range is selected when the script runs, it interferes
  // with script-driven selection. The fix is to clear this kind of selection.
  if (doc.selection && doc.selection.typename) {
    clearSelection();
  }

  // ================================================
  // grab custom settings, html, css, js and text blocks
  // ================================================
  var documentHasSettingsBlock = false;
  var customBlocks = {};
  var customRxp = /^ai2html-(css|js|html|settings|text)\s*$/;

  forEach(doc.textFrames, function(thisFrame) {
    // var contents = thisFrame.contents; // caused MRAP error in AI 2017
    var type = null;
    var match, entries;
    if (thisFrame.lines.length > 1) {
      match = customRxp.exec(thisFrame.lines[0].contents);
      type = match ? match[1] : null;
    }
    if (!type) return; // not a settings block
    entries = stringToLines(thisFrame.contents);
    entries.shift(); // remove header
    if (type == 'settings') {
      documentHasSettingsBlock = true;
      parseSettingsEntries(entries, docSettings);
    } else if (type == 'text') {
      parseSettingsEntries(entries, docSettings);
    } else { // custom js, css and html
      if (!customBlocks[type]) {
        customBlocks[type] = [];
      }
      customBlocks[type].push("\t\t" + cleanText(entries.join("\r\t\t")) + "\r");
    }
    if (objectOverlapsAnArtboard(thisFrame)) {
      hideTextFrame(thisFrame);
    }
  });

  if (customBlocks.css)  {feedback.push("Custom CSS blocks: " + customBlocks.css.length);}
  if (customBlocks.html) {feedback.push("Custom HTML blocks: " + customBlocks.html.length);}
  if (customBlocks.js)   {feedback.push("Custom JS blocks: " + customBlocks.js.length);}

  // ================================================
  // add settings text block if one does not exist
  // ================================================

  if (!documentHasSettingsBlock) {
    createSettingsBlock();
    if (scriptEnvironment=="nyt") {
      feedback.push("A settings text block was created to the left of all your artboards. Fill out the settings to link your project to the Scoop asset.");
      return; // Exit the script
    } else {
      feedback.push("A settings text block was created to the left of all your artboards. You can use it to customize your output.");
    }
  }


  // ================================================
  // assign artboards to their corresponding breakpoints
  // ================================================
  // (can have more than one artboard per breakpoint.)
  var breakpoints = assignBreakpointsToArtboards(nyt5Breakpoints);

  // ================================================
  // initialization for NYT environment
  // ================================================

  if (scriptEnvironment == "nyt") {
    // Read yml file to determine what type of project this is
    // (yml file should be confirmed to exist when nyt environment is set)
    var yaml = readYamlConfigFile(docPath + "../config.yml") || {};
    previewProjectType = yaml.project_type == 'ai2html' ? 'ai2html' : '';
    if ((previewProjectType=="ai2html" && !folderExists(docPath + "../public/")) ||
        (previewProjectType!="ai2html" && !folderExists(docPath + "../src/"))) {
      errors.push("Make sure your Illustrator file is inside the \u201Cai\u201D folder of a Preview project.");
      errors.push("If the Illustrator file is in the correct folder, your Preview project may be missing a \u201Cpublic\u201D or a \u201Csrc\u201D folder.");
      errors.push("If this is an ai2html project, it is probably easier to just create a new ai2html Preview project and move this Illustrator file into the \u201Cai\u201D folder inside the project.");
      return;
    }

    if (yaml.scoop_slug) {
      docSettings.scoop_slug_from_config_yml = yaml.scoop_slug;
    }
    // Read .git/config file to get preview slug
    var gitConfig = readGitConfigFile(docPath + "../.git/config") || {};
    if (gitConfig.url) {
      docSettings.preview_slug = gitConfig.url.replace( /^[^:]+:/ , "" ).replace( /\.git$/ , "");
    }

    docSettings.image_source_path = "_assets/";
    if (previewProjectType == "ai2html") {
      docSettings.html_output_path      = "/../public/";
      docSettings.html_output_extension = ".html";
      docSettings.image_output_path     = "_assets/";
    }

    if (docSettings.max_width && !contains(breakpoints, function(bp) {
      return +docSettings.max_width == bp.upperLimit;
    })) {
      warnings.push('The max_width setting of "' + docSettings.max_width +
        '" is not a valid breakpoint and will create an error when you "preview publish."');
    }
  }

  // ================================================
  // initialization for all environments
  // ================================================

  docName = docSettings.project_name || doc.name.replace(/(.+)\.[aieps]+$/,"$1").replace(/ +/g,"-");
  docName = makeKeyword(docName);

  if (docSettings.image_source_path === null) {
    docSettings.image_source_path = docSettings.image_output_path;
  }

  if (docSettings.image_format.length === 0) {
    warnings.push("No images were created because no image formats were specified.");
  } else if (contains(docSettings.image_format, "auto")) {
    docSettings.image_format = [documentContainsVisibleRasterImages() ? 'jpg' : 'png'];
  } else if (documentContainsVisibleRasterImages() && !contains(docSettings.image_format, "jpg")) {
    warnings.push("An artboard contains a raster image -- consider exporting to jpg instead of " +
        docSettings.image_format[0] + ".");
  }

  // ================================================
  // Generate HTML, CSS and images for each artboard
  // ================================================
  pBar = new ProgressBar({name: "Ai2html progress", steps: calcProgressBarSteps()});
  unlockObjects(); // Unlock containers and clipping masks
  var masks = findMasks(); // identify all clipping masks and their contents
  var artboardContent = {html: "", css: "", js: ""};

  forEachUsableArtboard(function(activeArtboard, abNumber) {
    var abSettings = getArtboardSettings(activeArtboard);
    var docArtboardName  = getArtboardFullName(activeArtboard);
    var textFrames, textData;
    doc.artboards.setActiveArtboardIndex(abNumber);

    // ========================
    // Convert text objects
    // ========================

    if (abSettings.image_only) {
      textFrames = [];
      textData = {html: "", styles: []};
    } else {
      pBar.setTitle(docArtboardName + ': Generating text...');
      textFrames = getTextFramesByArtboard(activeArtboard, masks);
      textData = convertTextFrames(textFrames, activeArtboard);
    }
    pBar.step();

    // ==========================
    // generate artboard image(s)
    // ==========================

    if (isTrue(docSettings.write_image_files)) {
      pBar.setTitle(docArtboardName + ': Capturing image...');
      captureArtboardImage(activeArtboard, textFrames, masks, docSettings);
    }
    pBar.step();

    //=====================================
    // finish generating artboard HTML and CSS
    //=====================================

    artboardContent.html += "\r\t<!-- Artboard: " + getArtboardName(activeArtboard) + " -->\r" +
       generateArtboardDiv(activeArtboard, breakpoints, docSettings) +
       generateImageHtml(activeArtboard, docSettings) +
       textData.html +
       "\t</div>\r";
    artboardContent.css += generateArtboardCss(activeArtboard, textData.styles, docSettings);
    /*
    artboardContent +=
      "\r\t<!-- Artboard: " + getArtboardName(activeArtboard) + " -->\r" +
      generateArtboardDiv(activeArtboard, breakpoints, docSettings) +
      generateArtboardCss(activeArtboard, textData.styles, docSettings) +
      generateImageHtml(activeArtboard, docSettings) +
      textData.html +
      "\t</div>\r";
    */

    //=====================================
    // output html file here if doing a file for every artboard
    //=====================================

    if (docSettings.output=="multiple-files") {
      addCustomContent(artboardContent, customBlocks);
      generateOutputHtml(artboardContent, docArtboardName, docSettings);
      artboardContent = {html: "", css: "", js: ""};
    }

  }); // end artboard loop

  //=====================================
  // output html file here if doing one file for all artboards
  //=====================================

  if (docSettings.output=="one-file") {
    addCustomContent(artboardContent, customBlocks);
    generateOutputHtml(artboardContent, docName, docSettings);
  }

  //=====================================
  // write configuration file with graphic metadata
  //=====================================

  if ((scriptEnvironment=="nyt" && previewProjectType=="ai2html") ||
      (scriptEnvironment!="nyt" && isTrue(docSettings.create_config_file))) {
    // TODO: switch to this?  (scriptEnvironment!="nyt" && docSettings.config_file_path)) {
    var yamlPath = docPath + docSettings.config_file_path,
        yamlStr = generateYamlFileContent(breakpoints, docSettings);
    checkForOutputFolder(yamlPath.replace(/[^\/]+$/, ""), "configFileFolder");
    saveTextFile(yamlPath, yamlStr);
  }

} // end render()


// =================================
// JS utility functions
// =================================

function forEach(arr, cb) {
  for (var i=0, n=arr.length; i<n; i++) {
    cb(arr[i], i);
  }
}

function map(arr, cb) {
  var arr2 = [];
  for (var i=0, n=arr.length; i<n; i++) {
    arr2.push(cb(arr[i], i));
  }
  return arr2;
}

function filter(arr, test) {
  var filtered = [];
  for (var i=0, n=arr.length; i<n; i++) {
    if (test(arr[i], i)) {
      filtered.push(arr[i]);
    }
  }
  return filtered;
}

// obj: value or test function
function indexOf(arr, obj) {
  var test = typeof obj == 'function' ? obj : null;
  for (var i=0, n=arr.length; i<n; i++) {
    if (test ? test(arr[i]) : arr[i] === obj) {
      return i;
    }
  }
  return -1;
}

function find(arr, obj) {
  var i = indexOf(arr, obj);
  return i == -1 ? null : arr[i];
}

function contains(arr, obj) {
  return indexOf(arr, obj) >= 0;
}

function extend(o) {
  for (var i=1; i<arguments.length; i++) {
    forEachProperty(arguments[i], add);
  }
  function add(v, k) {
    o[k] = v;
  }
  return o;
}

function forEachProperty(o, cb) {
  for (var k in o) {
    if (o.hasOwnProperty(k)) {
      cb(o[k], k);
    }
  }
}

// Return new object containing properties that are in a but not b
// Return null if output object would be empty
// a, b: JS objects
function objectSubtract(a, b) {
  var diff = null;
  for (var k in a) {
    if (a[k] != b[k] && a.hasOwnProperty(k)) {
      diff = diff || {};
      diff[k] = a[k];
    }
  }
  return diff;
}

// return elements in array "a" but not in array "b"
function arraySubtract(a, b) {
  var diff = [],
      alen = a.length,
      blen = b.length,
      i, j;
  for (i=0; i<alen; i++) {
    diff.push(a[i]);
    for (j=0; j<blen; j++) {
      if (a[i] === b[j]) {
        diff.pop();
        break;
      }
    }
  }
  return diff;
}

// Copy elements of an array-like object to an array
function toArray(obj) {
  var arr = [];
  for (var i=0, n=obj.length; i<n; i++) {
    arr[i] = obj[i];
  }
  return arr;
}

// multiple key sorting function based on https://github.com/Teun/thenBy.js
// first by length of name, then by population, then by ID
// data.sort(
//     firstBy(function (v1, v2) { return v1.name.length - v2.name.length; })
//     .thenBy(function (v1, v2) { return v1.population - v2.population; })
//     .thenBy(function (v1, v2) { return v1.id - v2.id; });
// );
function firstBy(f1, f2) {
  var compare = f2 ? function(a, b) {return f1(a, b) || f2(a, b);} : f1;
  compare.thenBy = function(f) {return firstBy(compare, f);};
  return compare;
}

function keys(obj) {
  var keys = [];
  for (var k in obj) {
    keys.push(k);
  }
  return keys;
}

// Remove whitespace from beginning and end of a string
function trim(s) {
  return s.replace(/^[\s\uFEFF\xA0\x03]+|[\s\uFEFF\xA0\x03]+$/g, '');
}

// splits a string into non-empty lines
function stringToLines(str) {
  var empty = /^\s*$/;
  return filter(str.split(/[\r\n\x03]+/), function(line) {
    return !empty.test(line);
  });
}

function zeroPad(val, digits) {
  var str = String(val);
  while (str.length < digits) str = '0' + str;
  return str;
}

function truncateString(str, maxlen) {
  // TODO: add ellipsis, truncate at word boundary
  if (str.length > maxlen) {
    str = str.substr(0, maxlen);
  }
  return str;
}

function makeKeyword(text) {
  return text.replace( /[^A-Za-z0-9_\-]+/g , "_" );
}

function cleanText(text) {
  for (var i=0; i < htmlCharacterCodes.length; i++) {
    var charCode = htmlCharacterCodes[i];
    if (text.indexOf(charCode[0]) > -1) {
      text = text.replace(new RegExp(charCode[0],'g'), charCode[1]);
    }
  }
  return text;
}

function straightenCurlyQuotesInsideAngleBrackets(text) {
  // This function's purpose is to fix quoted properties in HTML tags that were
  // typed into text blocks (Illustrator tends to automatically change single
  // and double quotes to curly quotes).
  // thanks to jashkenas
  var tagFinder = /<[^\n]+?>/g;
  var quoteFinder = /[\u201C‘’\u201D]([^\n]*?)[\u201C‘’\u201D]/g;
  return text.replace(tagFinder, function(tag){
    return tag.replace( /[\u201C\u201D]/g , '"' ).replace( /[‘’]/g , "'" );
  });
}

// Not very robust -- good enough for printing a warning
function findHtmlTag(str) {
  var match;
  if (str.indexOf('<') > -1) { // bypass regex check
    match = /<(\w+)[^>]*>/.exec(str);
  }
  return match ? match[1] : null;
}

// precision: number of decimals in rounded number
function roundTo(number, precision) {
  var d = Math.pow(10, precision || 0);
  return Math.round(number * d) / d;
}

function getDateTimeStamp() {
  var d     = new Date();
  var year  = d.getFullYear();
  var date  = zeroPad(d.getDate(),2);
  var month = zeroPad(d.getMonth() + 1,2);
  var hour  = zeroPad(d.getHours(),2);
  var min   = zeroPad(d.getMinutes(),2);
  return year + "-" + month + "-" + date + " " + hour + ":" + min;
}

// obj: JS object containing css properties and values
// indentStr: string to use as block CSS indentation
function formatCss(obj, indentStr) {
  var css = '';
  var isBlock = !!indentStr;
  for (var key in obj) {
    if (isBlock) {
      css += '\r' + indentStr;
    }
    css += key + ':' + obj[key]+ ';';
  }
  if (css && isBlock) {
    css += '\r';
  }
  return css;
}

function getCssColor(r, g, b, opacity) {
  var col, o;
  if (opacity > 0 && opacity < 100) {
    o = roundTo(opacity / 100, 2);
    col = 'rgba(' + r + ',' + g + ',' + b + ',' + o + ')';
  } else {
    col = 'rgb(' + r + ',' + g + ',' + b + ')';
  }
  return col;
}

// Test if two rectangles are the same, to within a given tolerance
// a, b: two arrays containing AI rectangle coordinates
// maxOffs: maximum pixel deviation on any side
function testSimilarBounds(a, b, maxOffs) {
  if (maxOffs >= 0 === false) maxOffs = 1;
  for (var i=0; i<4; i++) {
    if (Math.abs(a[i] - b[i]) > maxOffs) return false;
  }
  return true;
}

// Apply very basic string substitution to a template
function applyTemplate(template, replacements) {
  var keyExp = '([_a-zA-Z][\\w-]*)';
  var mustachePattern = new RegExp("\\{\\{\\{? *" + keyExp + " *\\}\\}\\}?","g");
  var ejsPattern = new RegExp("<%=? *" + keyExp + " *%>","g");
  var replace = function(match, name) {
    var lcname = name.toLowerCase();
    if (name in replacements) return replacements[name];
    if (lcname in replacements) return replacements[lcname];
    return match;
  };
  return template.replace(mustachePattern, replace).replace(ejsPattern, replace);
}



// ======================================
// Illustrator specific utility functions
// ======================================

// a, b: coordinate arrays, as from <PathItem>.geometricBounds
function testBoundsIntersection(a, b) {
  return a[2] >= b[0] && b[2] >= a[0] && a[3] <= b[1] && b[3] <= a[1];
}

function shiftBounds(bnds, dx, dy) {
  return [bnds[0] + dx, bnds[1] + dy, bnds[2] + dx, bnds[3] + dy];
}

function clearMatrixShift(m) {
  return app.concatenateTranslationMatrix(m, -m.mValueTX, -m.mValueTY);
}

function folderExists(path) {
  return new Folder(path).exists;
}

function fileExists(path) {
  return new File(path).exists;
}

function readYamlConfigFile(path) {
  return fileExists(path) ? parseYaml(readTextFile(path)) : null;
}

function parseKeyValueString(str, o) {
  var dqRxp = /^"(?:[^"\\]|\\.)*"$/;
  var parts = str.split(':');
  var k, v;
  if (parts.length > 1) {
    k = trim(parts.shift());
    v = trim(parts.join(':'));
    if (dqRxp.test(v)) {
      v = JSON.parse(v); // use JSON library to parse quoted strings
    }
    o[k] = v;
  }
}

// Very simple Yaml parsing. Does not implement nested properties and other features
function parseYaml(str) {
  // TODO: strip comments // var comment = /\s*/
  var o = {};
  var lines = stringToLines(str);
  for (var i = 0; i < lines.length; i++) {
    parseKeyValueString(lines[i], o);
  }
  return o;
}

// TODO: improve
// (currently ignores bracketed sections of the config file)
function readGitConfigFile(path) {
  var file = new File(path);
  var o = null;
  var parts;
  if (file.exists) {
    o = {};
    file.open("r");
    while(!file.eof) {
      parts = file.readln().split("=");
      if (parts.length > 1) {
        o[trim(parts[0])] = trim(parts[1]);
      }
    }
    file.close();
  }
  return o;
}

function readTextFile(path) {
  var outputText = "";
  var file = new File(path);
  if (file.exists) {
    file.open("r");
    while (!file.eof) {
      outputText += file.readln() + "\n";
    }
    file.close();
  } else {
    warnings.push(path + " could not be found.");
  }
  return outputText;
}

function saveTextFile(dest, contents) {
  var fd = new File(dest);
  fd.open("w", "TEXT", "TEXT");
  fd.lineFeed = "Unix";
  fd.encoding = "UTF-8";
  fd.writeln(contents);
  fd.close();
}

function checkForOutputFolder(folderPath, nickname) {
  var outputFolder = new Folder( folderPath );
  if (!outputFolder.exists) {
    var outputFolderCreated = outputFolder.create();
    if (outputFolderCreated) {
      feedback.push("The " + nickname + " folder did not exist, so the folder was created.");
    } else {
      warnings.push("The " + nickname + " folder did not exist and could not be created.");
    }
  }
}



// =====================================
// ai2html specific utility functions
// =====================================

function calcProgressBarSteps() {
  var n = 0;
  forEachUsableArtboard(function() {
    n += 2;
  });
  return n;
}

function formatError(e) {
  var msg = "Runtime error";
  if (e.line) msg += " on line " + e.line;
  if (e.message) msg += ": " + e.message;
  return msg;
}

function warnOnce(msg, item) {
  if (!contains(oneTimeWarnings, item)) {
    warnings.push(msg);
    oneTimeWarnings.push(item);
  }
}

// display debugging message in completion alert box
// (in debug mode)
function message() {
  var msg = "", arg;
  for (var i=0; i<arguments.length; i++) {
    arg = arguments[i];
    if (msg.length > 0) msg += ' ';
    if (typeof arg == 'object') {
      try {
        // json2.json implementation throws error if object contains a cycle
        // and many Illustrator objects have cycles.
        msg += JSON.stringify(arg);
      } catch(e) {
        msg += String(arg);
      }
    } else {
      msg += arg;
    }
  }
  if (showDebugMessages) feedback.push(msg);
}


// accept inconsistent true/yes setting value
function isTrue(val) {
  return val == "true" || val == "yes" || val === true;
}

// accept inconsistent false/no setting value
function isFalse(val) {
  return val == "false" || val == "no" || val === false;
}

function unlockObjects() {
  forEach(doc.layers, unlockContainer);
}

function unlockObject(obj) {
  obj.locked = false;
  objectsToRelock.push(obj);
}

// Unlock a layer or group if visible and locked, as well as any locked and visible
//   clipping masks
// o: GroupItem or Layer
function unlockContainer(o) {
  var type = o.typename;
  var i, item, pathCount;
  if (o.hidden === true || o.visible === false) return;
  if (o.locked) {
    unlockObject(o);
  }

  // unlock locked clipping paths (so contents can be selected later)
  // optimization: Layers containing hundreds or thousands of paths are unlikely
  //    to contain a clipping mask and are slow to scan -- skip these
  pathCount = o.pathItems.length;
  if ((type == 'Layer' && pathCount < 500) || (type == 'GroupItem' && o.clipped)) {
    for (i=0; i<pathCount; i++) {
      item = o.pathItems[i];
      if (!item.hidden && item.clipping && item.locked) {
        unlockObject(item);
        break;
      }
    }
  }

  // recursively unlock sub-layers and groups
  forEach(o.groupItems, unlockContainer);
  if (o.typename == 'Layer') {
    forEach(o.layers, unlockContainer);
  }
}



// ==================================
// ai2html program state and settings
// ==================================

function runningInNode() {
  return (typeof module != "undefined") && !!module.exports;
}

function isTestedIllustratorVersion(version) {
  var majorNum = parseInt(version);
  return majorNum >= 18 && majorNum <= 21; // Illustrator CC 2014 through 2017
}

function validateArtboardNames() {
  var names = [];
  forEachUsableArtboard(function(ab) {
    var name = getArtboardName(ab);
    if (contains(names, name)) {
      warnOnce("Artboards should have unique names. \"" + name + "\" is duplicated.", name);
    }
    names.push(name);
  });
}

function showEngineInfo() {
  var lines = map($.summary().split('\n'), function(line) {
    var parts = trim(line).split(/[\s]+/);
    if (parts.length == 2) {
      line = parts[1] + ' ' + parts[0];
    }
    return line;
  }).sort();
  var msg = lines.join('  ');
  // msg = $.list().split('\n').length;
  // msg = truncateString($.listLO(), 300);
  alert('Info:\n' + msg);
}

function detectScriptEnvironment() {
  var env = detectTimesFonts() ? 'nyt' : '';
  // Handle case where user seems to be at NYT but is running ai2html outside of Preview
  if (env == 'nyt' && !fileExists(docPath + "../config.yml")) {
    if(confirm("You seem to be running ai2html outside of NYT Preview.\nContinue in non-Preview mode?", true)) {
      env = ''; // switch to non-nyt context
    } else {
      errors.push("Ai2html should be run inside a Preview project.");
    }
  }
  return env;
}

function detectTimesFonts() {
  var found = false;
  try {
    app.textFonts.getByName('NYTFranklin-Medium') && app.textFonts.getByName('NYTCheltenham-Medium');
    found = true;
  } catch(e) {}
  return found;
}

function initDocumentSettings(env) {
  ai2htmlBaseSettings = env == 'nyt' ? nytBaseSettings : defaultBaseSettings;
  // initialize document settings
  docSettings = {};
  for (var setting in ai2htmlBaseSettings) {
    docSettings[setting] = ai2htmlBaseSettings[setting].defaultValue;
  }
}

function initUtilityFunctions() {
  // Enable timing using T.start() and T.stop("message")
  T = {
    stack: [],
    start: function() {
      T.stack.push(+new Date());
    },
    stop: function(note) {
      var ms = +new Date() - T.stack.pop();
      if (note) message(ms + 'ms - ' + note);
      return ms;
    }
  };
}

function createSettingsBlock() {
  var bounds      = getAllArtboardBounds();
  var fontSize    = 15;
  var leading     = 19;
  var extraLines  = 6;
  var width       = 400;
  var left        = bounds[0] - width - 50;
  var top         = bounds[1];
  var settingsLines = ["ai2html-settings"];
  var layer, rect, textArea, height;

  for (var name in ai2htmlBaseSettings) {
    if (ai2htmlBaseSettings[name].includeInSettingsBlock) {
      settingsLines.push(name + ": " + ai2htmlBaseSettings[name].defaultValue);
    }
  }

  try {
    layer = doc.layers.getByName("ai2html-settings");
  } catch(e) {
    layer = doc.layers.add();
    layer.zOrder(ZOrderMethod.BRINGTOFRONT);
    layer.name  = "ai2html-settings";
  }

  height = leading * (settingsLines.length + extraLines);
  rect = layer.pathItems.rectangle(top, left, width, height);
  textArea = layer.textFrames.areaText(rect);
  textArea.textRange.autoLeading = false;
  textArea.textRange.characterAttributes.leading = leading;
  textArea.textRange.characterAttributes.size = fontSize;
  textArea.contents = settingsLines.join('\n');
}


// Add ai2html settings from a text block to the document settings object
function parseSettingsEntries(entries, docSettings) {
  var entryRxp = /^([\w-]+)\s*:\s*(.*)$/;
  forEach(entries, function(str) {
    var match, hashKey, hashValue;
    str = trim(str);
    match = entryRxp.exec(str);
    if (!match) {
      if (str) warnings.push("Malformed setting, skipping: " + str);
      return;
    }
    hashKey   = match[1];
    hashValue = straightenCurlyQuotesInsideAngleBrackets(match[2]);
    if (hashKey in docSettings === false) {
      // assumes docSettings has been initialized with default settings
      warnings.push("Settings block contains an unsupported parameter: " + hashKey);
    }
    // replace values from old versions of script with current values
    if (hashKey=="output" && hashValue=="one-file-for-all-artboards") { hashValue="one-file"; }
    if (hashKey=="output" && hashValue=="one-file-per-artboard")      { hashValue="multiple-files"; }
    if (hashKey=="output" && hashValue=="preview-one-file")           { hashValue="one-file"; }
    if (hashKey=="output" && hashValue=="preview-multiple-files")     { hashValue="multiple-files"; }
    if ((hashKey in ai2htmlBaseSettings) && ai2htmlBaseSettings[hashKey].inputType=="array") {
      hashValue = hashValue.replace( /[\s,]+/g , ',' );
      if (hashValue.length === 0) {
        hashValue = []; // have to do this because .split always returns an array of length at least 1 even if it's splitting an empty string
      } else {
        hashValue = hashValue.split(",");
      }
    }
    docSettings[hashKey] = hashValue;
  });
}

// Show alert or prompt; return true if promo image should be generated
function showCompletionAlert(showPrompt) {
  var rule = "\n================\n";
  var alertText, alertHed, makePromo;

  if (errors.length > 0) {
    alertHed = "The Script Was Unable to Finish";
  } else if (scriptEnvironment == "nyt") {
    alertHed = "Actually, that\u2019s not half bad :)"; // &rsquo;
  } else {
    alertHed = "Nice work!";
  }
  alertText  = makeList(errors, "Error", "Errors");
  alertText += makeList(warnings, "Warning", "Warnings");
  alertText += makeList(feedback, "Information", "Information");
  alertText += "\n";
  if (showPrompt) {
    alertText += rule + "Generate promo image?";
    makePromo = confirm(alertHed  + alertText, true); // true: "No" is default
  } else {
    alertText += rule + "ai2html-nyt5 v" + scriptVersion;
    alert(alertHed + alertText);
    makePromo = false;
  }

  function makeList(items, singular, plural) {
    var list = "";
    if (items.length > 0) {
      list += "\r" + (items.length == 1 ? singular : plural) + rule;
      for (var i = 0; i < items.length; i++) {
        list += "\u2022 " + items[i] + "\r";
      }
    }
    return list;
  }
  return makePromo;
}

function restoreDocumentState() {
  var i;
  for (i = 0; i<textFramesToUnhide.length; i++) {
    textFramesToUnhide[i].hidden = false;
  }
  for (i = objectsToRelock.length-1; i>=0; i--) {
    objectsToRelock[i].locked = true;
  }
}

function ProgressBar(opts) {
  opts = opts || {};
  var steps = opts.steps || 0;
  var step = 0;
  var win = new Window("palette", opts.name || "Progress", [150, 150, 600, 260]);
  win.pnl = win.add("panel", [10, 10, 440, 100], "Progress");
  win.pnl.progBar      = win.pnl.add("progressbar", [20, 35, 410, 60], 0, 100);
  win.pnl.progBarLabel = win.pnl.add("statictext", [20, 20, 320, 35], "0%");
  win.show();

  function getProgress() {
    return win.pnl.progBar.value/win.pnl.progBar.maxvalue;
  }

  function update() {
    win.update();
  }

  this.step = function() {
    step = Math.min(step + 1, steps);
    this.setProgress(step / steps);
  };

  this.setProgress = function(progress) {
    var max = win.pnl.progBar.maxvalue;
    // progress is always 0.0 to 1.0
    var pct = progress * max;
    win.pnl.progBar.value = pct;
    win.pnl.progBarLabel.text = Math.round(pct) + "%";
    update();
  };

  this.setTitle = function(title) {
    win.pnl.text = title;
    update();
  };

  this.close = function() {
    win.close();
  };
}


// ======================================
// ai2html AI document reading functions
// ======================================

// Convert bounds coordinates (e.g. artboardRect, geometricBounds) to CSS-style coords
function convertAiBounds(rect) {
  var x = rect[0],
      y = -rect[1],
      w = Math.round(rect[2] - x),
      h = -rect[3] - y;
  return {
    left: x,
    top: y,
    width: w,
    height: h
  };
}

// Get numerical index of an artboard in the doc.artboards array
function getArtboardId(ab) {
  var id = 0;
  forEachUsableArtboard(function(ab2, i) {
    if (ab === ab2) id = i;
  });
  return id;
}

// TODO: prevent duplicate names? or treat duplicate names an an error condition?
// (artboard name is assumed to be unique in several places)
function getArtboardName(ab) {
  return makeKeyword(ab.name.replace( /^(.+):.*$/, "$1"));
}

function getArtboardFullName(ab) {
  return docName + "-" + getArtboardName(ab);
}

// return coordinates of bounding box of all artboards
function getAllArtboardBounds() {
  var rect, bounds;
  for (var i=0, n=doc.artboards.length; i<n; i++) {
    rect = doc.artboards[i].artboardRect;
    if (i === 0) {
      bounds = rect;
    } else {
      bounds = [
        Math.min(rect[0], bounds[0]), Math.max(rect[1], bounds[1]),
        Math.max(rect[2], bounds[2]), Math.min(rect[3], bounds[3])];
    }
  }
  return bounds;
}

// return responsive artboard widths as an array [minw, maxw]
function getArtboardWidthRange(ab) {
  var id = getArtboardId(ab);
  var infoArr = getArtboardInfo();
  var minw, maxw;
  // find min width, which is the artboard's own effective width
  forEach(infoArr, function(info) {
    if (info.id == id) {
      minw = info.effectiveWidth;
    }
  });
  // find max width, which is the effective width of the next widest
  // artboard (if any), minus one pixel
  forEach(infoArr, function(info) {
    var w = info.effectiveWidth;
    if (w > minw && (!maxw || w < maxw)) {
      maxw = w;
    }
  });
  return [minw, maxw ? maxw - 1 : Infinity];
}

// Parse artboard-specific settings from artboard name
function parseArtboardName(name) {
  // parse old-style width declaration
  var widthStr = (/^ai2html-(\d+)/.exec(name) || [])[1];
  // capture portion of name after colon
  var settingsStr = (/:(.*)/.exec(name) || [])[1] || "";
  var settings = {};
  forEach(settingsStr.split(','), function(part) {
    if (/^\d+$/.test(part)) {
      widthStr = part;
    } else if (part) {
      // assuming setting is a flag
      settings[part] = true;
    }
  });
  if (widthStr) {
    settings.width = parseFloat(widthStr);
  }
  return settings;
}

function getArtboardSettings(ab) {
  // currently, artboard-specific settings are all stashed in the artboard name
  return parseArtboardName(ab.name);
}

// return array of data records about each usable artboard, sorted from narrow to wide
function getArtboardInfo() {
  var artboards = [];
  forEachUsableArtboard(function(ab, i) {
    var pos = convertAiBounds(ab.artboardRect);
    var abSettings = getArtboardSettings(ab);
    artboards.push({
      name: ab.name || "",
      width: pos.width,
      effectiveWidth: abSettings.width || pos.width,
      id: i
    });
  });
  artboards.sort(function(a, b) {return a.effectiveWidth - b.effectiveWidth;});
  return artboards;
}

// Get array of data records for breakpoints that have artboards assigned to them
// (sorted from narrow to wide)
// breakpoints: Array of data about all possible breakpoints
function assignBreakpointsToArtboards(breakpoints) {
  var abArr = getArtboardInfo(); // get data records for each artboard
  var bpArr = [];
  forEach(breakpoints, function(breakpoint) {
    var bpPrev = bpArr[bpArr.length - 1],
        bpInfo = {
          name: breakpoint.name,
          lowerLimit: breakpoint.lowerLimit,
          upperLimit: breakpoint.upperLimit,
          artboards: []
        },
        abInfo;
    for (var i=0; i<abArr.length; i++) {
      abInfo = abArr[i];
      if (abInfo.effectiveWidth <= breakpoint.upperLimit &&
          abInfo.effectiveWidth > breakpoint.lowerLimit) {
        bpInfo.artboards.push(abInfo.id);
      }
    }
    if (bpInfo.artboards.length > 1 && scriptEnvironment=="nyt") {
      warnings.push('The ' + breakpoint.upperLimit + "px breakpoint has " + bpInfo.artboards.length +
          " artboards. You probably want only one artboard per breakpoint.");
    }
    if (bpInfo.artboards.length === 0 && bpPrev) {
      bpInfo.artboards = bpPrev.artboards.concat();
    }
    if (bpInfo.artboards.length > 0) {
      bpArr.push(bpInfo);
    }
  });
  return bpArr;
}

function forEachUsableArtboard(cb) {
  var ab;
  for (var i=0; i<doc.artboards.length; i++) {
    ab = doc.artboards[i];
    if (!/^-/.test(ab.name)) { // exclude artboards with names starting w/ "-"
      cb(ab, i);
    }
  }
}

// Returns id of artboard with largest area
function findLargestArtboard() {
  var largestId = -1;
  var largestArea = 0;
  forEachUsableArtboard(function(ab, i) {
    var info = convertAiBounds(ab.artboardRect);
    var area = info.width * info.height;
    if (area > largestArea) {
      largestId = i;
      largestArea = area;
    }
  });
  return largestId;
}

function clearSelection() {
  // setting selection to null doesn't always work:
  // it doesn't deselect text range selection and also seems to interfere with
  // subsequent mask operations using executeMenuCommand().
  // doc.selection = null;
  // the following seems to work reliably.
  app.executeMenuCommand('deselectall');
}

function documentContainsVisibleRasterImages() {
  // TODO: verify that placed items are rasters
  function isVisible(obj) {
    return !objectIsHidden(obj) && objectOverlapsAnArtboard(obj);
  }
  return contains(doc.placedItems, isVisible) || contains(doc.rasterItems, isVisible);
}

function objectOverlapsAnArtboard(obj) {
  var hit = false;
  forEachUsableArtboard(function(ab) {
    hit = hit || testBoundsIntersection(ab.artboardRect, obj.geometricBounds);
  });
  return hit;
}

function objectIsHidden(obj) {
  var hidden = false;
  while (!hidden && obj && obj.typename != "Document"){
    if (obj.typename == "Layer") {
      hidden = !obj.visible;
    } else {
      hidden = obj.hidden;
    }
    obj = obj.parent;
  }
  return hidden;
}

function getComputedOpacity(obj) {
  var opacity = 1;
  while (obj && obj.typename != "Document") {
    opacity *= obj.opacity / 100;
    obj = obj.parent;
  }
  return opacity * 100;
}


// Return array of layer objects, including both PageItems and sublayers, in z order
function getSortedLayerItems(lyr) {
  var items = toArray(lyr.pageItems).concat(toArray(lyr.layers));
  if (lyr.layers.length > 0 && lyr.pageItems.length > 0) {
    // only need to sort if layer contains both layers and page objects
    items.sort(function(a, b) {
      return b.absoluteZOrderPosition - a.absoluteZOrderPosition;
    });
  }
  return items;
}

// a, b: Layer objects
function findCommonLayer(a, b) {
  var p = null;
  if (a == b) {
    p = a;
  }
  if (!p && a.parent.typename == 'Layer') {
    p = findCommonLayer(a.parent, b);
  }
  if (!p && b.parent.typename == 'Layer') {
    p = findCommonLayer(a, b.parent);
  }
  return p;
}

function findCommonAncestorLayer(items) {
  var layers = [],
      ancestorLyr = null,
      item;
  for (var i=0, n=items.length; i<n; i++) {
    item = items[i];
    if (item.parent.typename != 'Layer' || contains(layers, item.parent)) {
      continue;
    }
    // remember layer, to avoid redundant searching (is this worthwhile?)
    layers.push(item.parent);
    if (!ancestorLyr) {
      ancestorLyr = item.parent;
    } else {
      ancestorLyr = findCommonLayer(ancestorLyr, item.parent);
      if (!ancestorLyr) {
        // Failed to find a common ancestor
        return null;
      }
    }
  }
  return ancestorLyr;
}

// Test if a mask can be ignored
// (An optimization -- currently only finds group masks with no text frames)
function maskIsRelevant(mask) {
  var parent = mask.parent;
  if (parent.typename == "GroupItem") {
    if (parent.textFrames.length === 0) {
      return false;
    }
  }
  return true;
}

function findMasks() {
  var found = [],
      masks, relevantMasks;
  // assumes clipping paths have been unlocked
  app.executeMenuCommand('Clipping Masks menu item');
  masks = toArray(doc.selection);
  clearSelection();
  relevantMasks = filter(masks, maskIsRelevant);
  forEach(masks, function(mask) {mask.locked = true;});
  forEach(relevantMasks, function(mask) {
    var items, obj;
    mask.locked = false;
    // select a single mask
    // some time ago, executeMenuCommand() was more reliable than assigning to
    // selection... no longer, apparently
    // app.executeMenuCommand('Clipping Masks menu item');
    doc.selection = [mask];
    // switch selection to all masked items
    app.executeMenuCommand('editMask'); // Object > Clipping Mask > Edit Contents
    items = toArray(doc.selection || []);
    // oddly, 'deselectall' sometimes fails here
    // app.executeMenuCommand('deselectall');
    doc.selection = null;
    mask.locked = true;
    obj = {
      mask: mask,
      items: items
    };
    if (mask.parent.typename == "GroupItem") {
      // Group mask
      obj.group = mask.parent;

    } else if (mask.parent.typename == "Layer") {
      // Layer mask -- common ancestor layer of all masked items is assumed
      // to be the masked layer
      obj.layer = findCommonAncestorLayer(items);

    } else {
      message("Unknown mask type in findMasks()");
    }

    if (items.length > 0 && (obj.group || obj.layer)) {
      found.push(obj);
    }

    if (items.length === 0) {
      // message("Unable to select masked items");
    }
  });
  forEach(masks, function(mask) {mask.locked = false;});
  return found;
}



// ==============================
// ai2html text functions
// ==============================

function textIsTransformed(textFrame) {
  var m = textFrame.matrix;
  return !(m.mValueA == 1 && m.mValueB === 0 && m.mValueC === 0 && m.mValueD == 1);
}

function hideTextFrame(textFrame) {
  textFramesToUnhide.push(textFrame);
  textFrame.hidden = true;
}

// color: a color object, e.g. RGBColor
// opacity (optional): opacity [0-100]
function convertAiColor(color, opacity) {
  var o = {};
  var r, g, b;
  if (color.typename == 'SpotColor') {
    color = color.spot.color; // expecting AI to return an RGBColor because doc is in RGB mode.
  }
  if (color.typename == 'RGBColor') {
    r = color.red;
    g = color.green;
    b = color.blue;
    if (r < rgbBlackThreshold && g < rgbBlackThreshold && b < rgbBlackThreshold) {
      r = g = b = 0;
    }
  } else if (color.typename == 'GrayColor') {
    r = g = b = Math.round((100 - color.gray) / 100 * 255);
  } else if (color.typename == 'NoColor') {
    g = 255;
    r = b = 0;
    // warnings are processed later, after ranges of same-style chars are identified
    // TODO: add text-fill-specific warnings elsewhere
    o.warning = "The text \"%s\" has no fill. Please fill it with an RGB color. It has been filled with green.";
  } else {
    r = g = b = 0;
    o.warning = "The text \"%s\" has " + color.typename + " fill. Please fill it with an RGB color.";
  }
  o.color = getCssColor(r, g, b, opacity);
  return o;
}

// Parse an AI CharacterAttributes object
function getCharStyle(c) {
  var o = convertAiColor(c.fillColor);
  var caps = String(c.capitalization);
  o.aifont = c.textFont.name;
  o.size = Math.round(c.size);
  o.capitalization = caps == 'FontCapsOption.NORMALCAPS' ? '' : caps;
  o.tracking = c.tracking
  return o;
}

// p: an AI paragraph (appears to be a TextRange object with mixed-in ParagraphAttributes)
// opacity: Computed opacity (0-100) of TextFrame containing this pg
function getParagraphStyle(p) {
  return {
    leading: Math.round(p.leading),
    spaceBefore: Math.round(p.spaceBefore),
    spaceAfter: Math.round(p.spaceAfter),
    justification: String(p.justification) // coerce from object
  };
}

// s: object containing CSS text properties
function getStyleKey(s) {
  var key = '';
  for (var i=0, n=cssTextStyleProperties.length; i<n; i++) {
    key += '~' + (s[cssTextStyleProperties[i]] || '');
  }
  return key;
}

function getTextStyleClass(style, classes, name) {
  var key = getStyleKey(style);
  var cname = nameSpace + (name || 'style');
  var o, i;
  for (i=0; i<classes.length; i++) {
    o = classes[i];
    if (o.key == key) {
      return o.classname;
    }
  }
  o = {
    key: key,
    style: style,
    classname: cname + i
  };
  classes.push(o);
  return o.classname;
}

// Divide a paragraph (TextRange object) into an array of
// data objects describing text strings having the same style.
function getParagraphRanges(p) {
  var segments = [];
  var currRange;
  var prev, curr, c;
  for (var i=0, n=p.characters.length; i<n; i++) {
    c = p.characters[i];
    curr = getCharStyle(c);
    if (!prev || objectSubtract(curr, prev)) {
      currRange = {
        text: "",
        aiStyle: curr
      };
      segments.push(currRange);
    }
    if (curr.warning) {
      currRange.warning = curr.warning;
    }
    currRange.text += c.contents;
    prev = curr;
  }
  return segments;
}


// Convert a TextFrame to an array of data records for each of the paragraphs
//   contained in the TextFrame.
function importTextFrameParagraphs(textFrame) {
  // The scripting API doesn't give us access to opacity of TextRange objects
  //   (including individual characters). The best we can do is get the
  //   computed opacity of the current TextFrame
  var opacity = getComputedOpacity(textFrame);
  var blendMode = getBlendMode(textFrame);
  var charsLeft = textFrame.characters.length;
  var data = [];
  var p, plen, d;
  for (var k=0, n=textFrame.paragraphs.length; k<n && charsLeft > 0; k++) {
    // trailing newline in a text block adds one to paragraphs.length, but
    // an error is thrown when such a pg is accessed. charsLeft test is a workaround.
    p = textFrame.paragraphs[k];
    plen = p.characters.length;
    if (plen === 0) {
      d = {
        text: "",
        aiStyle: {},
        ranges: []
      };
    } else {
      d = {
        text: p.contents,
        aiStyle: getParagraphStyle(p, opacity),
        ranges: getParagraphRanges(p)
      };
      d.aiStyle.opacity = opacity;
      d.aiStyle.blendMode = blendMode;
    }
    data.push(d);
    charsLeft -= (plen + 1); // char count + newline
  }
  return data;
}

function cleanHtmlTags(str) {
  var tagName = findHtmlTag(str);
  // only warn for certain tags
  if (tagName && contains('i,span,b,strong,em'.split(','), tagName.toLowerCase())) {
    warnOnce("Found a <" + tagName + "> tag. Try using Illustrator formatting instead.", tagName);
  }
  return tagName ? straightenCurlyQuotesInsideAngleBrackets(str) : str;
}

function generateParagraphHtml(pData, baseStyle, pStyles, cStyles) {
  var html, diff, classname, range, text;
  if (pData.text.length === 0) { // empty pg
    // TODO: Calculate the height of empty paragraphs and generate
    // CSS to preserve this height (not supported by Illustrator API)
    return '<p>&nbsp;</p>';
  }
  diff = objectSubtract(pData.cssStyle, baseStyle);
  // Give the pg a class, if it has a different style than the base pg class
  if (diff) {
    classname = getTextStyleClass(diff, pStyles, 'pstyle');
    html = '<p class="' + classname + '">';
  } else {
    html = '<p>';
  }
  for (var j=0; j<pData.ranges.length; j++) {
    range = pData.ranges[j];
    range.text = cleanHtmlTags(range.text);
    diff = objectSubtract(range.cssStyle, pData.cssStyle);
    if (diff) {
      classname = getTextStyleClass(diff, cStyles, 'cstyle');
      html += '<span class="' + classname + '">';
    }
    html += cleanText(range.text);
    if (diff) {
      html += '</span>';
    }
  }
  html += '</p>';
  return html;
}

function generateTextFrameHtml(paragraphs, baseStyle, pStyles, cStyles) {
  var html = "";
  for (var i=0; i<paragraphs.length; i++) {
    html += '\r\t\t\t' + generateParagraphHtml(paragraphs[i], baseStyle, pStyles, cStyles);
  }
  return html;
}

// Convert a collection of TextFrames to HTML and CSS
function convertTextFrames(textFrames, ab) {
  var frameData = map(textFrames, function(frame, i) {
    return {
      paragraphs: importTextFrameParagraphs(frame)
    };
  });
  var pgStyles = [];
  var charStyles = [];
  var baseStyle = deriveCssStyles(frameData);
  var idPrefix = nameSpace + "ai" + getArtboardId(ab) + "-";
  var abBox = convertAiBounds(ab.artboardRect);
  var divs = map(frameData, function(obj, i) {
    var frame = textFrames[i];
    var divId = frame.name ? makeKeyword(frame.name) : idPrefix  + (i + 1);
    var positionCss = getTextFrameCss(frame, abBox, obj.paragraphs);
    return '\t\t<div id="' + divId + '" ' + positionCss + '>' +
        generateTextFrameHtml(obj.paragraphs, baseStyle, pgStyles, charStyles) + '\r\t\t</div>\r';
  });

  var allStyles = pgStyles.concat(charStyles);
  var cssBlocks = map(allStyles, function(obj) {
    return '.' + obj.classname + ' {' + formatCss(obj.style, '\t\t') + '\t}\r';
  });
  if (divs.length > 0) {
    cssBlocks.unshift('p {' + formatCss(baseStyle, '\t\t') + '\t}\r');
  }

  return {
    styles: cssBlocks,
    html: divs.join('')
  };
}

// Compute the base paragraph style by finding the most common style in frameData
// Side effect: adds cssStyle object alongside each aiStyle object
// frameData: Array of data objects parsed from a collection of TextFrames
// Returns object containing css text style properties of base pg style
function deriveCssStyles(frameData) {
  var pgStyles = [];
  var baseStyle = {};
  // override detected settings with these style properties
  var defaultCssStyle = {
    'text-align': 'left',
    'text-transform': 'none',
    'padding-bottom': 0,
    'padding-top': 0,
    'mix-blend-mode': 'normal',
    'font-style': 'normal'
  };
  var defaultAiStyle = {
    opacity: 100 // given as AI style because opacity is converted to several CSS properties
  };
  var currCharStyles;

  forEach(frameData, function(frame) {
    forEach(frame.paragraphs, analyzeParagraphStyle);
  });

  // find the most common pg style and override certain properties
  if (pgStyles.length > 0) {
    pgStyles.sort(compareCharCount);
    extend(baseStyle, pgStyles[0].cssStyle);
  }
  extend(baseStyle, defaultCssStyle, convertAiTextStyle(defaultAiStyle));
  return baseStyle;

  function compareCharCount(a, b) {
    return b.count - a.count;
  }

  function analyzeParagraphStyle(pdata) {
    currCharStyles = [];
    forEach(pdata.ranges, convertRangeStyle);
    if (currCharStyles.length > 0) {
      // add most common char style to the pg style, to avoid applying
      // <span> tags to all the text in the paragraph
      currCharStyles.sort(compareCharCount);
      extend(pdata.aiStyle, currCharStyles[0].aiStyle);
    }
    pdata.cssStyle = analyzeTextStyle(pdata.aiStyle, pdata.text, pgStyles);
    if (pdata.aiStyle.blendMode && !pdata.cssStyle['mix-blend-mode']) {
      warnOnce("Missing a rule for converting " + pdata.aiStyle.blendMode + " to CSS.", pdata.aiStyle.blendMode);
    }
  }

  function convertRangeStyle(range) {
    range.cssStyle = analyzeTextStyle(range.aiStyle, range.text, currCharStyles);
    if (range.warning) {
      warnings.push(range.warning.replace("%s", truncateString(range.text, 35)));
    }
    if (range.aiStyle.aifont && !range.cssStyle['font-family']) {
      warnOnce("Missing a rule for converting font: " + range.aiStyle.aifont +
        ". Sample text: " + truncateString(range.text, 35), range.aiStyle.aifont);
    }
  }

  function analyzeTextStyle(aiStyle, text, stylesArr) {
    var cssStyle = convertAiTextStyle(aiStyle);
    var key = getStyleKey(cssStyle);
    var o;
    if (text.length === 0) {
      return {};
    }
    for (var i=0; i<stylesArr.length; i++) {
      if (stylesArr[i].key == key) {
        o = stylesArr[i];
        break;
      }
    }
    if (!o) {
      o = {
        key: key,
        aiStyle: aiStyle,
        cssStyle: cssStyle,
        count: 0
      };
      stylesArr.push(o);
    }
    o.count += text.length;
    // o.count++; // each occurence counts equally
    return cssStyle;
  }
}


// Lookup an AI font name in the font table
function findFontInfo(aifont) {
  var info = null;
  for (var k=0; k<fonts.length; k++) {
    if (aifont == fonts[k].aifont) {
      info = fonts[k];
      break;
    }
  }
  if (!info) {
    // font not found... parse the AI font name to give it a weight and style
    info = {};
    if (aifont.indexOf('Italic') > -1) {
      info.style = 'italic';
    }
    if (aifont.indexOf('Bold') > -1) {
      info.weight = 700;
    } else {
      info.weight = 500;
    }
  }
  return info;
}

// ai: AI justification value
function getJustificationCss(ai) {
  for (var k=0; k<align.length; k++) {
    if (ai == align[k].ai) {
      return align[k].html;
    }
  }
  return "initial"; // CSS default
}

// ai: AI capitalization value
function getCapitalizationCss(ai) {
  for (var k=0; k<caps.length; k++) {
    if (ai == caps[k].ai) {
      return caps[k].html;
    }
  }
  return "";
}

function getBlendModeCss(ai) {
  for (var k=0; k<blendModes.length; k++) {
    if (ai == blendModes[k].ai) {
      return blendModes[k].html;
    }
  }
  return "";
}

function getBlendMode(obj) {
  // Limitation: returns first found blending mode, ignores any others that
  //   might be applied a parent object
  while (obj && obj.typename != "Document") {
    if (obj.blendingMode && obj.blendingMode != BlendModes.NORMAL) {
      return obj.blendingMode;
    }
    obj = obj.parent;
  }
  return null;
}

// convert an object containing parsed AI text styles to an object containing CSS style properties
function convertAiTextStyle(aiStyle) {
  var cssStyle = {};
  var fontInfo, tmp;
  if (aiStyle.aifont) {
    fontInfo = findFontInfo(aiStyle.aifont);
    if (fontInfo.family) {
      cssStyle["font-family"] = fontInfo.family;
    }
    if (fontInfo.weight) {
      cssStyle["font-weight"] = fontInfo.weight;
    }
    if (fontInfo.style) {
      cssStyle["font-style"] = fontInfo.style;
    }
  }
  if (aiStyle.size > 0) {
    cssStyle["font-size"] = aiStyle.size + "px";
  }
  if ('leading' in aiStyle) {
    cssStyle["line-height"] = aiStyle.leading + "px";
  }
  // if (('opacity' in aiStyle) && aiStyle.opacity < 100) {
  if ('opacity' in aiStyle) {
    cssStyle.filter = "alpha(opacity=" + Math.round(aiStyle.opacity) + ")";
    cssStyle["-ms-filter"] = "progid:DXImageTransform.Microsoft.Alpha(Opacity=" +
        Math.round(aiStyle.opacity) + ")";
    cssStyle.opacity = roundTo(aiStyle.opacity / 100, cssPrecision);
  }
  if (aiStyle.blendMode && (tmp = getBlendModeCss(aiStyle.blendMode))) {
    cssStyle['mix-blend-mode'] = tmp;
    // TODO: consider opacity fallback for IE
  }
  if (aiStyle.spaceBefore > 0) {
    cssStyle["padding-top"] = aiStyle.spaceBefore + "px";
  }
  if (aiStyle.spaceAfter > 0) {
    cssStyle["padding-bottom"] = aiStyle.spaceAfter + "px";
  }
  if ('tracking' in aiStyle) {
    cssStyle["letter-spacing"] = roundTo(aiStyle.tracking / 1000, cssPrecision) + "em";
  }
  if (aiStyle.justification && (tmp = getJustificationCss(aiStyle.justification))) {
    cssStyle["text-align"] = tmp;
  }
  if (aiStyle.capitalization && (tmp = getCapitalizationCss(aiStyle.capitalization))) {
    cssStyle["text-transform"] = tmp;
  }
  if (aiStyle.color) {
    cssStyle.color = aiStyle.color;
  }
  return cssStyle;
}

function textFrameIsRenderable(frame, artboardRect) {
  var good = true;
  if (!testBoundsIntersection(frame.visibleBounds, artboardRect)) {
    good = false;
  } else if (frame.kind != TextType.AREATEXT && frame.kind != TextType.POINTTEXT) {
    good = false;
  } else if (objectIsHidden(frame)) {
    good = false;
  } else if (frame.contents === "") {
    good = false;
  } else if (docSettings.render_rotated_skewed_text_as == "image" && textIsTransformed(frame)) {
    good = false;
  }
  return good;
}

// Find clipped art objects that are inside an artboard but outside the bounding box
// box of their clipping path
// items: array of PageItems assocated with a clipping path
// clipRect: bounding box of clipping path
// abRect: bounds of artboard to test
//
function selectMaskedItems(items, clipRect, abRect) {
  var found = [];
  var itemRect, itemInArtboard, itemInMask, maskInArtboard;
  for (var i=0; i<items.length; i++) {
    itemRect = items[i].geometricBounds;
    // capture items that intersect the artboard but are masked...
    itemInArtboard = testBoundsIntersection(abRect, itemRect);
    maskInArtboard = testBoundsIntersection(abRect, clipRect);
    itemInMask = testBoundsIntersection(itemRect, clipRect);
    if (itemInArtboard && (!maskInArtboard || !itemInMask)) {
      found.push(items[i]);
    }
  }
  return found;
}

// Find clipped TextFrames that are inside an artboard but outside their
// clipping path (using bounding box of clipping path to approximate clip area)
function getClippedTextFramesByArtboard(ab, masks) {
  var abRect = ab.artboardRect;
  var frames = [];
  forEach(masks, function(o) {
    var clipRect = o.mask.geometricBounds;
    if (testSimilarBounds(abRect, clipRect, 5)) {
      // if clip path is masking the current artboard, skip the test
      // (optimization)
      return;
    }
    var texts = filter(o.items, function(item) {return item.typename == 'TextFrame';});
    texts = selectMaskedItems(texts, clipRect, abRect);
    if (texts.length > 0) {
      frames = frames.concat(texts);
    }
  });
  return frames;
}

// Get array of TextFrames belonging to an artboard, excluding text that
// overlaps the artboard but is hidden by a clipping mask
function getTextFramesByArtboard(ab, masks) {
  var candidateFrames = findTextFramesToRender(doc.textFrames, ab.artboardRect);
  var excludedFrames = getClippedTextFramesByArtboard(ab, masks);
  var goodFrames = arraySubtract(candidateFrames, excludedFrames);
  return goodFrames;
}

function findTextFramesToRender(frames, artboardRect) {
  var selected = [];
  for (var i=0; i<frames.length; i++) {
    if (textFrameIsRenderable(frames[i], artboardRect)) {
      selected.push(frames[i]);
    }
  }
  // Sort frames top to bottom, left to right.
  selected.sort(
      firstBy(function (v1, v2) { return v2.top  - v1.top; })
      .thenBy(function (v1, v2) { return v1.left - v2.left; })
  );
  return selected;
}

// Extract key: value pairs from the contents of a note attribute
function parseDataAttributes(note) {
  var o = {};
  var parts, part;
  if (note) {
    parts = note.split(/[\r\n;,]+/);
    for (var i = 0; i < parts.length; i++) {
      parseKeyValueString(parts[i], o);
    }
  }
  return o;
}

function formatCssPct(part, whole) {
  return roundTo(part / whole * 100, cssPrecision) + "%;";
}

function getUntransformedTextBounds(textFrame) {
  var copy = textFrame.duplicate(textFrame.parent, ElementPlacement.PLACEATEND);
  var matrix = clearMatrixShift(textFrame.matrix);
  copy.transform(app.invertMatrix(matrix));
  var bnds = copy.geometricBounds;
  if (textFrame.kind == TextType.AREATEXT) {
    // prevent offcenter problem caused by extra vertical space in text area
    // TODO: de-kludge
    // this would be much simpler if <TextFrameItem>.convertAreaObjectToPointObject()
    // worked correctly (throws MRAP error when trying to remove a converted object)
    var textWidth = (bnds[2] - bnds[0]);
    copy.transform(matrix);
    // Transforming outlines avoids the offcenter problem, but width of bounding
    // box needs to be set to width of transformed TextFrame for correct output
    copy = copy.createOutline();
    copy.transform(app.invertMatrix(matrix));
    bnds = copy.geometricBounds;
    var dx = Math.ceil(textWidth - (bnds[2] - bnds[0])) / 2;
    bnds[0] -= dx;
    bnds[2] += dx;
  }
  copy.remove();
  return bnds;
}

function getTransformationCss(textFrame, vertAnchorPct) {
  var matrix = clearMatrixShift(textFrame.matrix);
  var horizAnchorPct = 50;
  var transformOrigin = horizAnchorPct + '% ' + vertAnchorPct + '%;';
  var transform = "matrix(" +
      roundTo(matrix.mValueA, cssPrecision) + ',' +
      roundTo(-matrix.mValueB, cssPrecision) + ',' +
      roundTo(-matrix.mValueC, cssPrecision) + ',' +
      roundTo(matrix.mValueD, cssPrecision) + ',' +
      roundTo(matrix.mValueTX, cssPrecision) + ',' +
      roundTo(matrix.mValueTY, cssPrecision) + ');';

  // TODO: handle character scaling.
  // One option: add separate CSS transform to paragraphs inside a TextFrame
  var charStyle = textFrame.textRange.characterAttributes;
  var scaleX = charStyle.horizontalScale;
  var scaleY = charStyle.verticalScale;
  if (scaleX != 100 || scaleY != 100) {
    warnings.push("Vertical or horizontal text scaling will be lost. Affected text: " + truncateString(textFrame.contents, 35));
  }

  return "transform: " + transform +  "transform-origin: " + transformOrigin +
    "-webkit-transform: " + transform + "-webkit-transform-origin: " + transformOrigin +
    "-ms-transform: " + transform + "-ms-transform-origin: " + transformOrigin;
}

// Create class="" and style="" CSS for positioning a text div
function getTextFrameCss(thisFrame, abBox, pgData) {
  var styles = "";
  var classes = "";
  var isTransformed = textIsTransformed(thisFrame);
  var aiBounds = isTransformed ? getUntransformedTextBounds(thisFrame) : thisFrame.geometricBounds;
  var htmlBox = convertAiBounds(shiftBounds(aiBounds, -abBox.left, abBox.top));
  var thisFrameAttributes = parseDataAttributes(thisFrame.note);
  // Using AI style of first paragraph in TextFrame to get information about
  // tracking, justification and top padding
  // TODO: consider positioning paragraphs separately, to handle pgs with different
  //   justification in the same text block
  var firstPgStyle = pgData[0].aiStyle;
  var lastPgStyle = pgData[pgData.length - 1].aiStyle;
  // estimated space between top of HTML container and character glyphs
  // (related to differences in AI and CSS vertical positioning of text blocks)
  var marginTopPx = (firstPgStyle.leading - firstPgStyle.size) / 2 + firstPgStyle.spaceBefore;
  // estimated space between bottom of HTML container and character glyphs
  var marginBottomPx = (lastPgStyle.leading - lastPgStyle.size) / 2 + lastPgStyle.spaceAfter;
  var trackingPx = firstPgStyle.size * firstPgStyle.tracking / 1000;
  var htmlL = htmlBox.left;
  var htmlT = Math.round(htmlBox.top - marginTopPx);
  var htmlW = htmlBox.width;
  var htmlH = htmlBox.height + marginTopPx + marginBottomPx;
  var alignment, v_align, vertAnchorPct;

  if (firstPgStyle.justification == "Justification.LEFT") {
    alignment = "left";
  } else if (firstPgStyle.justification == "Justification.RIGHT") {
    alignment = "right";
  } else if (firstPgStyle.justification == "Justification.CENTER") {
    alignment = "center";
  }

  if (isTransformed) {
    vertAnchorPct = (marginTopPx + htmlBox.height * 0.5 + 1) / (htmlH) * 100; // TODO: de-kludge
    styles += getTransformationCss(thisFrame, vertAnchorPct);
  }

  if (thisFrame.kind == TextType.AREATEXT) {
    v_align = "top"; // area text aligned to top by default
    // EXPERIMENTAL feature
    // Put a box around the text, if the text frame's textPath is styled
    styles += convertAreaTextPath(thisFrame);
  } else {
    // point text aligned to midline (sensible default for chart y-axes, map labels, etc.)
    v_align = "middle";
  }

  if (thisFrameAttributes.valign) {
    // override default vertical alignment
    v_align = thisFrameAttributes.valign;
    if (v_align == "center") {
      v_align = "middle";
    }
  }

  if (v_align == "bottom") {
    var bottomPx = abBox.height - (htmlBox.top + htmlBox.height + marginBottomPx);
    styles += "bottom:" + formatCssPct(bottomPx, abBox.height);
  } else if (v_align == "middle") {
    // https://css-tricks.com/centering-in-the-unknown/
    // TODO: consider: http://zerosixthree.se/vertical-align-anything-with-just-3-lines-of-css/
    styles += "top:" + formatCssPct(htmlT + marginTopPx + htmlBox.height / 2, abBox.height);
    styles += "margin-top:-" + roundTo(marginTopPx + htmlBox.height / 2, 1) + 'px;';
  } else {
    styles += "top:" + formatCssPct(htmlT, abBox.height);
  }
  if (alignment == "right") {
    styles += "right:" + formatCssPct(abBox.width - (htmlL + htmlBox.width), abBox.width);
  } else if (alignment == "center") {
    styles += "left:" + formatCssPct(htmlL + htmlBox.width/ 2, abBox.width);
    styles += "margin-left:" + formatCssPct(-htmlW / 2, abBox.width);
  } else {
    styles += "left:" + formatCssPct(htmlL, abBox.width);
  }

  classes = nameSpace + makeKeyword(thisFrame.layer.name) + " " + nameSpace + "aiAbs";
  if (thisFrame.kind == TextType.POINTTEXT) {
    classes += ' ' + nameSpace + 'aiPointText';
    // using pixel width with point text, because pct width causes alignment problems -- see issue #63
    // adding extra pixels in case HTML width is slightly less than AI width (affects alignment of right-aligned text)
    styles += "width:" + roundTo(htmlW + 2, cssPrecision) + 'px;';
  } else {
    // area text uses pct width, so width of text boxes will scale
    // TODO: consider only using pct width with wider text boxes that contain paragraphs of text
    styles += "width:" + formatCssPct(htmlW, abBox.width);
  }
  return 'class="' + classes + '" style="' + styles + '"';
}


function convertAreaTextPath(frame) {
  var style = "";
  var path = frame.textPath;
  var obj;
  if (path.stroked || path.filled) {
    style += "padding: 6px 6px 6px 7px;";
    if (path.filled) {
      obj = convertAiColor(path.fillColor, path.opacity);
      style += "background-color: " + obj.color + ";";
    }
    if (path.stroked) {
      obj = convertAiColor(path.strokeColor, path.opacity);
      style += "border: 1px solid " + obj.color + ";";
    }
  }
  return style;
}


// =================================
// ai2html image functions
// =================================

// ab: artboard (assumed to be the active artboard)
// textFrames:  text frames belonging to the active artboard
function captureArtboardImage(ab, textFrames, masks, settings) {
  var docArtboardName = getArtboardFullName(ab);
  var imageDestinationFolder = docPath + settings.html_output_path + settings.image_output_path;
  var imageDestination = imageDestinationFolder + docArtboardName;
  var i;
  checkForOutputFolder(imageDestinationFolder, "image_output_path");

  if (!isTrue(settings.testing_mode)) {
    for (i=0; i<textFrames.length; i++) {
      textFrames[i].hidden = true;
    }
  }

  exportImageFiles(imageDestination, ab, settings.image_format, 1, docSettings.use_2x_images_if_possible);
  if (contains(settings.image_format, 'svg')) {
    exportSVG(imageDestination, ab, masks);
  }

  if (!isTrue(settings.testing_mode)) {
    for (i=0; i<textFrames.length; i++) {
      textFrames[i].hidden = false;
    }
  }
}

// Create an <img> tag for the artboard image
function generateImageHtml(ab, settings) {
  var abName = getArtboardFullName(ab),
      abPos = convertAiBounds(ab.artboardRect),
      imgId = nameSpace + "ai" + getArtboardId(ab) + "-0",
      extension = (settings.image_format[0] || "png").substring(0,3),
      src = settings.image_source_path + abName + "." + extension,
      html;

  html = '\t\t<img id="' + imgId + '" class="' + nameSpace + 'aiImg"';
  if (isTrue(settings.use_lazy_loader)) {
    html += ' data-src="' + src + '"';
    // spaceholder while image loads
    src = 'data:image/gif;base64,R0lGODlhCgAKAIAAAB8fHwAAACH5BAEAAAAALAAAAAAKAAoAAAIIhI+py+0PYysAOw==';
  }
  html += ' src="' + src + '"/>\r';
  return html;
}

// Create a promo image from the largest usable artboard
function createPromoImage(settings) {
  var PROMO_WIDTH = 1024;
  var abNumber = findLargestArtboard();
  if (abNumber == -1) return; // TODO: show error

  var artboard         =  doc.artboards[abNumber],
      abPos            =  convertAiBounds(artboard.artboardRect),
      promoScale       =  PROMO_WIDTH / abPos.width,
      promoW           =  abPos.width * promoScale,
      promoH           =  abPos.height * promoScale,
      imageDestination =  docPath + docName + "-promo",
      promoFormat, tmpPngTransparency;

  // Previous file name was more complicated:
  // imageDestination = docPath + docSettings.docName + "-" + makeKeyword(ab.name) + "-" + abNumber + "-promo";

  doc.artboards.setActiveArtboardIndex(abNumber);

  // Using "jpg" if present in image_format setting, else using "png";
  if (contains(settings.image_format, 'jpg')) {
    promoFormat = 'jpg';
  } else {
    promoFormat = 'png';
  }

  tmpPngTransparency = settings.png_transparent;
  settings.png_transparent = "no";
  exportImageFiles(imageDestination, artboard, [promoFormat], promoScale, "no");
  settings.png_transparent = tmpPngTransparency;
  alert("Promo image created\nLocation: " + imageDestination + "." + promoFormat);
}

// Returns 1 or 2 (corresponding to standard pixel scale and "retina" pixel scale)
// format: png, png24 or jpg
// doubleres: yes, always or no (no is default)
function getOutputImagePixelRatio(width, height, format, doubleres) {
  // Maximum pixel sizes are based on mobile Safari limits
  // TODO: check to see if these numbers are still relevant
  var maxPngSize = 3*1024*1024;
  var maxJpgSize = 32*1024*1024;
  var k = (doubleres == "always" || doubleres == "yes") ? 2 : 1;
  var pixels = width * height * k * k;

  if (doubleres == "yes" && width < 945) { // assume wide images are desktop-only
    // use single res if image might run into mobile browser limits
    if (((format == "png" || format == "png24") && pixels > maxPngSize) ||
        (format == "jpg" && pixels > maxJpgSize)) {
      k = 1;
    }
  }
  return k;
}

// Exports contents of active artboard as an image (without text, unless in test mode)
//
// dest: full path of output file excluding the file extension
// ab: assumed to be active artboard
// formats: array of export format identifiers (png, png24, jpg)
// initialScaling: the proportion to scale the base image before considering whether to double res. Usually just 1.
// doubleres: "yes", "no" or "always" ("yes" may be overridden if the image is very large)
//
function exportImageFiles(dest, ab, formats, initialScaling, doubleres) {

  forEach(formats, function(format) {
    var maxJpgScale  = 776.19; // This is specified in the Illustrator Scripting Reference under ExportOptionsJPEG.
    var abPos = convertAiBounds(ab.artboardRect);
    var width = abPos.width * initialScaling;
    var height = abPos.height * initialScaling;
    var imageScale = 100 * initialScaling * getOutputImagePixelRatio(width, height, format, doubleres);
    var exportOptions, fileType;

    if (format=="png") {
      fileType = ExportType.PNG8;
      exportOptions = new ExportOptionsPNG8();
      exportOptions.colorCount       = docSettings.png_number_of_colors;
      exportOptions.transparency     = isTrue(docSettings.png_transparent);

    } else if (format=="png24") {
      fileType = ExportType.PNG24;
      exportOptions = new ExportOptionsPNG24();
      exportOptions.transparency     = isTrue(docSettings.png_transparent);

    } else if (format=="jpg") {
      if (imageScale > maxJpgScale) {
        imageScale = maxJpgScale;
        warnings.push(dest.split("/").pop() + ".jpg was output at a smaller size than desired because of a limit on jpg exports in Illustrator." +
          " If the file needs to be larger, change the image format to png which does not appear to have limits.");
      }
      fileType = ExportType.JPEG;
      exportOptions = new ExportOptionsJPEG();
      exportOptions.qualitySetting = docSettings.jpg_quality;

    } else {
      if (format != "svg") { // svg exported separately
        warnings.push("Unsupported image format: " + format);
      }
      return;
    }

    exportOptions.horizontalScale  = imageScale;
    exportOptions.verticalScale    = imageScale;
    exportOptions.artBoardClipping = true;
    exportOptions.antiAliasing     = false;
    app.activeDocument.exportFile(new File(dest), fileType, exportOptions);
  });
}


// Copy contents of an artboard to a temporary document, excluding objects
// that are hidden by masks
// TODO: grouped text is copied (but hidden). Avoid copying text in groups, for
//   smaller SVG output.
function copyArtboardForImageExport(ab, masks) {
  var layerMasks = filter(masks, function(o) {return !!o.layer;}),
      artboardBounds = ab.artboardRect,
      sourceLayers = toArray(doc.layers),
      destLayer = doc.layers.add(),
      destGroup = doc.groupItems.add(),
      groupPos, group2, doc2;

  destLayer.name = "ArtboardContent";
  destGroup.move(destLayer, ElementPlacement.PLACEATEND);
  forEach(sourceLayers, copyLayer);
  // need to save group position before copying to second document. Oddly,
  // the reported position of the original group changes after duplication
  groupPos = destGroup.position;
  // create temp document (pretty slow -- ~1.5s)
  doc2 = app.documents.add(DocumentColorSpace.RGB, doc.width, doc.height, 1);
  doc2.pageOrigin = doc.pageOrigin; // not sure if needed
  doc2.rulerOrigin = doc.rulerOrigin;
  doc2.artboards[0].artboardRect = artboardBounds;
  group2 = destGroup.duplicate(doc2.layers[0], ElementPlacement.PLACEATEND);
  group2.position = groupPos;
  destGroup.remove();
  destLayer.remove();
  return doc2;

  function copyLayer(lyr) {
    var mask;
    if (lyr.hidden) return; // ignore hidden layers
    mask = findLayerMask(lyr);
    if (mask) {
      copyMaskedLayerAsGroup(lyr, mask);
    } else {
      forEach(getSortedLayerItems(lyr), copyLayerItem);
    }
  }

  function removeHiddenItems(group) {
    // only remove text frames, for performance
    // TODO: consider checking all item types
    // TODO: consider checking subgroups (recursively)
    // FIX: convert group.textFrames to array to avoid runtime error "No such element" in forEach()
    forEach(toArray(group.textFrames), removeItemIfHidden);
  }

  function removeItemIfHidden(item) {
    if (item.hidden) item.remove();
  }

  // Item: Layer (sublayer) or PageItem
  function copyLayerItem(item) {
    if (item.typename == 'Layer') {
      copyLayer(item);
    } else {
      copyPageItem(item, destGroup);
    }
  }

  // TODO: locked objects in masked layer may not be included in mask.items array
  //   consider traversing layer in this function ...
  //   make sure doubly masked objects aren't copied twice
  function copyMaskedLayerAsGroup(lyr, mask) {
    var maskBounds = mask.mask.geometricBounds;
    var newMask, newGroup;
    if (!testBoundsIntersection(artboardBounds, maskBounds)) {
      return;
    }
    newGroup = doc.groupItems.add();
    newGroup.move(destGroup, ElementPlacement.PLACEATEND);
    forEach(mask.items, function(item) {
      copyPageItem(item, newGroup);
    });
    if (newGroup.pageItems.length > 0) {
      // newMask = duplicateItem(mask.mask, destGroup);
      // TODO: refactor
      newMask = mask.mask.duplicate(destGroup, ElementPlacement.PLACEATEND);
      newMask.moveToBeginning(newGroup);
      newGroup.clipped = true;
    } else {
      newGroup.remove();
    }
  }

  function findLayerMask(lyr) {
    return find(layerMasks, function(o) {return o.layer == lyr;});
  }

  function copyPageItem(item, dest) {
    var excluded =
        // item.typename == 'TextFrame' || // text objects should be copied if visible
        !testBoundsIntersection(item.geometricBounds, artboardBounds) ||
        objectIsHidden(item) || item.clipping;
    var copy;
    if (!excluded) {
      copy = item.duplicate(dest, ElementPlacement.PLACEATEND); //  duplicateItem(item, dest);
      if (copy.typename == 'GroupItem') {
        removeHiddenItems(copy);
      }
    }
  }
}

function exportSVG(dest, ab, masks) {
  // Illustrator's SVG output contains all objects in a document (it doesn't
  //   clip to the current artboard), so we copy artboard objects to a temporary
  //   document for export.
  var exportDoc = copyArtboardForImageExport(ab, masks);
  var opts = new ExportOptionsSVG();
  opts.embedAllFonts         = false;
  opts.fontSubsetting        = SVGFontSubsetting.None;
  opts.compressed            = false;
  opts.documentEncoding      = SVGDocumentEncoding.UTF8;
  opts.embedRasterImages     = isTrue(docSettings.svg_embed_images);
  opts.DTD                   = SVGDTDVersion.SVG1_1;
  opts.cssProperties         = SVGCSSPropertyLocation.STYLEATTRIBUTES;

  exportDoc.exportFile(new File(dest), ExportType.SVG, opts);
  doc.activate();
  //exportDoc.pageItems.removeAll();
  exportDoc.close(SaveOptions.DONOTSAVECHANGES);
}


// ===================================
// ai2html output generation functions
// ===================================

function generateArtboardDiv(ab, breakpoints, settings) {
  var divId = nameSpace + getArtboardFullName(ab);
  var classnames = nameSpace + "artboard " + nameSpace + "artboard-v3";
  var widthRange = getArtboardWidthRange(ab);
  var html = "";
  if (!isFalse(settings.include_resizer_classes)) {
    classnames += " " + findShowClassesForArtboard(ab, breakpoints);
  }
  html += '\t<div id="' + divId + '" class="' + classnames + '"';
  if (isTrue(settings.include_resizer_widths)) {
    // add data-min/max-width attributes
    // TODO: see if we can use breakpoint data to set min and max widths
    html += " data-min-width='" + widthRange[0] + "'";
    if (widthRange[1] < Infinity) {
      html += " data-max-width='" + widthRange[1] + "'";
    }
  }
  html += ">\r";
  return html;
}

function findShowClassesForArtboard(ab, breakpoints) {
  var classes = [];
  var id = getArtboardId(ab);
  forEach(breakpoints, function(bp) {
    if (contains(bp.artboards, id)) {
      classes.push(nameSpace + 'show-' + bp.name);
    }
  });
  return classes.join(' ');
}

function generateArtboardCss(ab, textClasses, settings) {
  var t3 = '\t',
      t4 = t3 + '\t',
      abId = "#" + nameSpace + getArtboardFullName(ab),
      css = "";
  css += t3 + abId + " {\r";
  css += t4 + "position:relative;\r";
  css += t4 + "overflow:hidden;\r";
  if (settings.responsiveness=="fixed") {
    css += t4 + "width:"  + convertAiBounds(ab.artboardRect).width + "px;\r";
  }
  css += t3 + "}\r";

  // classes for paragraph and character styles
  forEach(textClasses, function(cssBlock) {
    css += t3 + abId + " " + cssBlock;
  });
  return css;
}

// Get CSS styles that are common to all generated content
function generatePageCss(containerId, settings) {
  var css = "";
  var t2 = '\t';
  var t3 = '\t\t';

  if (!!settings.max_width) {
    css += t2 + "#" + containerId + " {\r";
    css += t3 + "max-width:" + settings.max_width + "px;\r";
    css += t2 + "}\r";
  }
  if (isTrue(settings.center_html_output)) {
    css += t2 + "#" + containerId + " ." + nameSpace + "artboard {\r";
    css += t3 + "margin:0 auto;\r";
    css += t2 + "}\r";
  }
  if (settings.clickable_link !== "") {
    css += t2 + "." + nameSpace + "ai2htmlLink {\r";
    css += t3 + "display: block;\r";
    css += t2 + "}\r";
  }
  // default <p> styles
  css += t2 + "#" + containerId + " ." + nameSpace + "artboard p {\r";
  css += t3 + "margin:0;\r";
  if (isTrue(settings.testing_mode)) {
    css += t3 + "color: rgba(209, 0, 0, 0.5) !important;\r";
  }
  css += t2 + "}\r";

  css += t2 + "." + nameSpace + "aiAbs {\r";
  css += t3 + "position:absolute;\r";
  css += t2 + "}\r";

  css += t2 + "." + nameSpace + "aiImg {\r";
  css += t3 + "display:block;\r";
  css += t3 + "width:100% !important;\r";
  css += t2 + "}\r";

  css += t2 + '.' + nameSpace + 'aiPointText p { white-space: nowrap; }\r';
  return css;
}

function generateYamlFileContent(breakpoints, settings) {
  var lines = [];
  lines.push("ai2html_version: " + scriptVersion);
  lines.push("project_type: " + previewProjectType);
  lines.push("tags: ai2html");
  lines.push("min_width: " + breakpoints[0].upperLimit); // TODO: ask why upperLimit
  if (!!settings.max_width) {
    lines.push("max_width: " + settings.max_width);
  } else if (settings.responsiveness != "fixed" && scriptEnvironment == "nyt") {
    lines.push("max_width: " + breakpoints[breakpoints.length-1].upperLimit);
  } else if (settings.responsiveness != "fixed" && scriptEnvironment != "nyt") {
    // don't write a max_width setting as there should be no max width in this case
  } else {
    // this is the case of fixed responsiveness
    lines.push("max_width: " + getArtboardInfo().pop().effectiveWidth);
  }
  return lines.join('\n') + '\n' + convertSettingsToYaml(settings) + '\n';
}

function convertSettingsToYaml(settings) {
  var lines = [];
  var value, useQuotes;
  for (var setting in settings) {
    if ((setting in ai2htmlBaseSettings) && ai2htmlBaseSettings[setting].includeInConfigFile) {
      value = trim(String(settings[setting]));
      useQuotes = value === "" || /\s/.test(value);
      if (setting == "show_in_compatible_apps") {
        // special case: this setting takes quoted "yes" or "no"
        useQuotes = true; // assuming value is 'yes' or 'no';
      }
      if (useQuotes) {
        value = JSON.stringify(value); // wrap in quotes and escape internal quotes
      } else if (isTrue(value) || isFalse(value)) {
        // use standard values for boolean settings
        value = isTrue(value) ? "true" : "false";
      }
      lines.push(setting + ': ' + value);
    }
  }
  return lines.join('\n');
}

function getResizerScript() {
  // The resizer function is embedded in the HTML page -- external variables must
  // be passed in.
  var resizer = function (scriptEnvironment, nameSpace) {
    // only want one resizer on the page
    if (document.documentElement.className.indexOf(nameSpace + "resizer-v3-init") > -1) return;
    document.documentElement.className += " " + nameSpace + "resizer-v3-init";
    // require IE9+
    if (!("querySelector" in document)) return;
    function updateSize() {
      var elements = Array.prototype.slice.call(document.querySelectorAll("." + nameSpace + "artboard-v3[data-min-width]")),
          widthById = {};
      elements.forEach(function(el) {
        var parent = el.parentNode,
            width = widthById[parent.id] || Math.round(parent.getBoundingClientRect().width),
            minwidth = el.getAttribute("data-min-width"),
            maxwidth = el.getAttribute("data-max-width");
        if (parent.id) widthById[parent.id] = width; // only if parent.id is set
        if (+minwidth <= width && (+maxwidth >= width || maxwidth === null)) {
          var img = el.querySelector("." + nameSpace + "aiImg");
          if (img.getAttribute("data-src") && img.getAttribute("src") != img.getAttribute("data-src")) {
            img.setAttribute("src", img.getAttribute("data-src"));
          }
          el.style.display = "block";
        } else {
          el.style.display = "none";
        }
      });

      if (scriptEnvironment=="nyt") {
        try {
          if (window.parent && window.parent.$) {
            window.parent.$("body").trigger("resizedcontent", [window]);
          }
          document.documentElement.dispatchEvent(new Event("resizedcontent"));
          if (window.require && document.querySelector("meta[name=sourceApp]") && document.querySelector("meta[name=sourceApp]").content == "nyt-v5") {
            require(["foundation/main"], function() {
              require(["shared/interactive/instances/app-communicator"], function(AppCommunicator) {
                AppCommunicator.triggerResize();
              });
            });
          }
        } catch(e) { console.log(e); }
      }
    }

    updateSize();

    window.addEventListener('nyt:embed:load', updateSize); // for nyt vi compatibility
    document.addEventListener("DOMContentLoaded", updateSize);

    window.addEventListener("resize", throttle(updateSize, 200));

    // based on underscore.js
    function throttle(func, wait) {
      var _now = Date.now || function() { return +new Date(); },
          timeout = null, previous = 0;
      var run = function() {
          previous = _now();
          timeout = null;
          func();
      };
      return function() {
        var remaining = wait - (_now() - previous);
        if (remaining <= 0 || remaining > wait) {
          if (timeout) {
            clearTimeout(timeout);
          }
          run();
        } else if (!timeout) {
          timeout = setTimeout(run, remaining);
        }
      };
    }
  };

  // convert function to JS source code
  var resizerJs = '(' +
    trim(resizer.toString().replace(/  /g, '\t')) + // indent with tabs
    ')("' + scriptEnvironment + '", "' + nameSpace + '");';
  return '<script type="text/javascript">\r\t' + resizerJs + '\r</script>\r';
}


// Write an HTML page to a file for NYT Preview
function outputLocalPreviewPage(textForFile, localPreviewDestination, settings) {
  var localPreviewTemplateText = readTextFile(docPath + settings.local_preview_template);
  settings.ai2htmlPartial = textForFile; // TODO: don't modify global settings this way
  var localPreviewHtml = applyTemplate(localPreviewTemplateText, settings);
  saveTextFile(localPreviewDestination, localPreviewHtml);
}

function addCustomContent(content, customBlocks) {
  if (customBlocks.css) {
    content.css += "\r\t\t/* Custom CSS */\r\t\t" + customBlocks.css.join('\r\t\t') + '\r';
    /*
    content = "\r\t<style type='text/css' media='screen,print'>\r" +
      "\t\t" + customBlocks.css.join('\r\t\t') +
      "\t</style>\r" + content;
    */
  }
  if (customBlocks.html) {
    content.html += "\r\t<!-- Custom HTML -->\r" + customBlocks.html.join('\r') + '\r';
  }
  // TODO: assumed JS contained in <script> tag -- verify this?
  if (customBlocks.js) {
    content.js += "\r\t<!-- Custom JS -->\r" + customBlocks.js.join('\r') + '\r';
  }
}

// Wrap content HTML in a <div>, add styles and resizer script, write to a file
function generateOutputHtml(content, pageName, settings) {
  var linkSrc = settings.clickable_link || "";
  var responsiveCss = "";
  var responsiveJs = "";
  var containerId = nameSpace + pageName + "-box";
  var textForFile, html, js, css, commentBlock;
  var htmlFileDestinationFolder;

  pBar.setTitle('Writing HTML output...');

  if (scriptEnvironment == "nyt" && !isFalse(settings.include_resizer_css_js)) {
    responsiveJs = '\t<script src="_assets/resizerScript.js"></script>' + "\n";
    if (previewProjectType == "ai2html") {
      responsiveCss = '\t<link rel="stylesheet" href="_assets/resizerStyle.css">' + "\n";
    }
  }
  if (isTrue(settings.include_resizer_script)) {
    responsiveJs  = getResizerScript();
    responsiveCss = "";
  }

  // comments
  commentBlock = "<!-- Generated by ai2html v" + scriptVersion + " - " +
    getDateTimeStamp() + " -->\r" + "<!-- ai file: " + doc.name + " -->\r";

  if (scriptEnvironment == "nyt") {
    commentBlock += "<!-- preview: " + settings.preview_slug + " -->\r";
  }
  if (settings.scoop_slug_from_config_yml) {
    commentBlock += "<!-- scoop: " + settings.scoop_slug_from_config_yml + " -->\r";
  }

  // HTML
  html = '<div id="' + containerId + '" class="ai2html">\r';
  if (linkSrc) {
    // optional link around content
    html += "\t<a class='" + nameSpace + "ai2htmlLink' href='" + linkSrc + "'>\r";
  }
  html += content.html;
  if (linkSrc) {
    html += "\t</a>\r";
  }
  html += "\r</div>\r";

  // CSS
  css = "<style type='text/css' media='screen,print'>\r" +
    generatePageCss(containerId, settings) +
    content.css +
    "\r</style>\r" + responsiveCss;

  // JS
  js = content.js + responsiveJs;

  if (scriptEnvironment == "nyt") {
    html = '<!-- SCOOP HTML -->\r' + commentBlock + html;
    css = '<!-- SCOOP CSS -->\r' + commentBlock + css;
    if (js) js ='<!-- SCOOP JS -->\r' + commentBlock + js;
  }

  textForFile = css + '\r' + html + '\r' + js;

  if (scriptEnvironment != "nyt") {
    textForFile = commentBlock + textForFile +
        "<!-- End ai2html" + " - " + getDateTimeStamp() + " -->\r";
  }

  textForFile = applyTemplate(textForFile, settings);
  htmlFileDestinationFolder = docPath + settings.html_output_path;
  checkForOutputFolder(htmlFileDestinationFolder, "html_output_path");
  htmlFileDestination = htmlFileDestinationFolder + pageName + settings.html_output_extension;

  if (settings.output == 'one-file' && previewProjectType == 'ai2html') {
    htmlFileDestination = htmlFileDestinationFolder + "index" + settings.html_output_extension;
  }

  // write file
  saveTextFile(htmlFileDestination, textForFile);

  // process local preview template if appropriate
  if (settings.local_preview_template !== "") {
    // TODO: may have missed a condition, need to compare with original version
    var previewFileDestination = htmlFileDestinationFolder + pageName + ".preview.html";
    outputLocalPreviewPage(textForFile, previewFileDestination, settings);
  }
}
