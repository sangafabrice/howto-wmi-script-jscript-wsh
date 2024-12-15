/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * @version 0.0.1.69
 */

// #region: header of utils.js
// Constants and variables.

/** @constant */
var BUTTONS_OKONLY = 0;
/** @constant */
var POPUP_ERROR = 16;
/** @constant */
var POPUP_NORMAL = 0;

/** @typedef */
var FileSystemObject = new ActiveXObject('Scripting.FileSystemObject');
/** @typedef */
var WshShell = new ActiveXObject('WScript.Shell');

var ScriptRoot = FileSystemObject.GetParentFolderName(WSH.ScriptFullName)

// #endregion

// #region: main
// The main part of the program.

/** @typedef */
var Package = getPackage();
/** @typedef */
var Param = getParameters();

/** The application execution. */
if (Param.Markdown && Package.IconLink.IsValid()) {
  /** @constant */
  var CMD_LINE_FORMAT = '"{0}" "{1}"';
  if (run(format(CMD_LINE_FORMAT, Package.IconLink.Path, Param.Markdown))) {
    popup('An unhandled error has occurred.', POPUP_ERROR);
  }
  quit(0);
}

/** Configuration and settings. */
if (Param.Set ^ Param.Unset) {
  // #region: setup.js
  // Methods for managing the shortcut menu option: install and uninstall.

  /** @typedef */
  var Setup = (function() {
    var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
    var KEY_FORMAT = 'HKCU\\{0}\\';
    var VERBICON_VALUENAME;

    return {
      /** Configure the shortcut menu in the registry. */
      Set: function () {
        VERB_KEY = format(KEY_FORMAT, VERB_KEY);
        var COMMAND_KEY = VERB_KEY + 'command\\';
        var command = format('{0} "{1}" /Markdown:"%1"', WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WSH.ScriptFullName);
        WshShell.RegWrite(COMMAND_KEY, command);
        WshShell.RegWrite(VERB_KEY, 'Convert to &HTML');
        VERBICON_VALUENAME = VERB_KEY + '\\Icon';
      },

      /**
       * Add an icon to the shortcut menu in the registry.
       * @param {string} menuIconPath is the shortcut menu icon file path.
       */
      AddIcon: function (menuIconPath) {
        WshShell.RegWrite(VERBICON_VALUENAME, menuIconPath);
      },

      /** Remove the shortcut icon menu. */
      RemoveIcon: function () {
        try {
          WshShell.RegDelete(VERBICON_VALUENAME);
        } catch (e) { }
      },

      /** Remove the shortcut menu by removing the verb key and subkeys. */
      Unset: function () {
        /** @typedef */
        var StdRegProv = GetObject('winmgmts:StdRegProv');
        var stdRegProvMethods = StdRegProv.Methods_;
        var enumKeyMethod = stdRegProvMethods('EnumKey');
        var enumKeyMethodParams = enumKeyMethod.InParameters;
        var inParams = enumKeyMethodParams.SpawnInstance_();
        inParams.hDefKey = 0x80000001; // HKCU
        // Recursion is used because a key with subkeys cannot be deleted.
        // Recursion helps removing the leaf keys first.
        (function(key) {
          inParams.sSubKeyName = key;
          var outParams = StdRegProv.ExecMethod_(enumKeyMethod.Name, inParams);
          var sNames = outParams.sNames;
          outParams = null;
          if (sNames != null) {
            var sNamesArray = sNames.toArray();
            for (var index = 0; index < sNamesArray.length; index++) {
              arguments.callee(format('{0}\\{1}', key, sNamesArray[index]));
            }
          }
          try {
            WshShell.RegDelete(format(KEY_FORMAT, key));
          } catch (e) { }
        })(VERB_KEY);
        inParams = null;
        enumKeyMethodParams = null;
        enumKeyMethod = null;
        stdRegProvMethods = null;
        StdRegProv = null;
      }
    }
  })();

  // #endregion

  if (Param.Set) {
    Package.IconLink.Create();
    Setup.Set();
    if (Param.NoIcon) {
      Setup.RemoveIcon();
    } else {
      Setup.AddIcon(Package.MenuIconPath);
    }
  } else if (Param.Unset) {
    Setup.Unset();
    Package.IconLink.Delete();
  }
  quit(0);
}

quit(1);

// #endregion

// #region: utils.js
// Utility functions.

/**
 * Generate a random file path.
 * @param {string} extension is the file extension.
 * @returns {string} a random file path.
 */
function generateRandomPath(extension) {
  var typeLib = new ActiveXObject('Scriptlet.TypeLib');
  try {
    return FileSystemObject.BuildPath(WshShell.ExpandEnvironmentStrings('%TEMP%'), typeLib.Guid.substr(1, 36).toLowerCase() + '.tmp' + extension);
  } finally {
    typeLib = null;
  }
}

/**
 * Delete the specified file.
 * @param {string} filePath is the file path.
 */
function deleteFile(filePath) {
  try {
    FileSystemObject.DeleteFile(filePath);
  } catch (e) { }
}

/**
 * Show the application message box.
 * @param {string} messageText is the message text to show.
 * @param {number} [popupType = POPUP_NORMAL] is the type of popup box.
 * @param {number} [popupButtons = BUTTONS_OKONLY] are the buttons of the message box.
 */
function popup(messageText, popupType, popupButtons) {
  if (!popupType) {
    popupType = POPUP_NORMAL;
  }
  if (!popupButtons) {
    popupButtons = BUTTONS_OKONLY;
  }
  run(format('"{0}" """"{1}"""" {2} {3}', Package.MessageBoxLinkPath, messageText.replace(/"/g, "'"), popupButtons, popupType));
}

/**
 * Run a command.
 * @param {string} commandLine
 * @returns {number} the exit status.
 */
function run(commandLine) {
  /** @constant */
  var WINDOW_STYLE_HIDDEN = 0;
  /** @constant */
  var WAIT_ON_RETURN = true;
  return WshShell.Run(commandLine, WINDOW_STYLE_HIDDEN, WAIT_ON_RETURN);
}

/**
 * Replace the format item "{n}" by the nth input in a list of arguments.
 * @param {string} formatStr the pattern format.
 * @param {...string} args the replacement texts.
 * @returns {string} a copy of format with the format items replaced by args.
 */
function format(formatStr, args) {
  args = Array.prototype.slice.call(arguments).slice(1);
  while (args.length > 0) {
    formatStr = formatStr.replace(new RegExp('\\{' + (args.length - 1) + '\\}', 'g'), args.pop());
  }
  return formatStr;
}

/** Destroy the COM objects. */
function dispose() {
  WshShell = null;
  FileSystemObject = null;
}

/**
 * Clean up and quit.
 * @param {number} exitCode .
 */
function quit(exitCode) {
  dispose();
  WSH.Quit(exitCode);
}

// #endregion

// #region: package.js
// Information about the resource files used by the project.

/** Get the package type. */
function getPackage() {
  /** @constant */
  var POWERSHELL_SUBKEY = 'HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe\\';
  /** The project resources directory path. */
  var resourcePath = FileSystemObject.BuildPath(ScriptRoot, 'rsc');

  /** The powershell core runtime path. */
  var pwshExePath = WshShell.RegRead(POWERSHELL_SUBKEY);
  /** The shortcut target powershell script path. */
  var pwshScriptPath = FileSystemObject.BuildPath(resourcePath, 'cvmd2html.ps1');
  var msgBoxLinkPath = FileSystemObject.BuildPath(ScriptRoot, 'MsgBox.lnk');
  var menuIconPath = FileSystemObject.BuildPath(resourcePath, 'menu.ico');

  /**
   * Store the partial "arguments" property string of the custom icon link.
   * The command is partial because it does not include the markdown file path string.
   * The markdown file path string will be input when calling the shortcut link.
   */
  var iconLinkArguments = format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown', pwshScriptPath);

  return {
    /** The shortcut menu icon path. */
    MenuIconPath: menuIconPath,
    /** The message box link path. */
    MessageBoxLinkPath: msgBoxLinkPath,

    /** Represents an adapted link object. */
    IconLink: {
      /** The icon link path. */
      Path: FileSystemObject.BuildPath(WshShell.SpecialFolders('StartMenu'), 'cvmd2html.lnk'),

      /** Create the custom icon link file. */
      Create: function () {
        var link = WshShell.CreateShortcut(this.Path);
        link.TargetPath = pwshExePath;
        link.Arguments = iconLinkArguments;
        link.IconLocation = menuIconPath;
        link.Save();
        link = null;
      },

      /** Delete the custom icon link file. */
      Delete: function () {
        deleteFile(this.Path);
      },

      /**
       * Validate the link properties.
       * @returns {boolean} true if the link properties are as expected, false otherwise.
       */
      IsValid: function () {
        var linkItem = WshShell.CreateShortcut(this.Path);
        var targetCommand = '{0} {1}';
        try {
          return format(targetCommand, linkItem.TargetPath, linkItem.Arguments).toLowerCase() == format(targetCommand, pwshExePath, iconLinkArguments).toLowerCase();
        } finally {
          linkItem = null;
        }
      }
    }
  }
}

// #endregion

// #region: parameters.js
// Parsed parameters.

/**
 * @typedef {object} ParamHash
 * @property {string} Markdown is the selected markdown file path.
 * @property {boolean} Set installs the shortcut menu.
 * @property {boolean} NoIcon installs the shortcut menu without icon.
 * @property {boolean} Unset uninstalls the shortcut menu.
 * @property {boolean} Help shows help.
 */

/** @returns {ParamHash} */
function getParameters() {
  var WshArguments = WSH.Arguments;
  var WshNamed = WshArguments.Named;
  var paramCount = WshArguments.Count();
  if (paramCount == 1) {
    var paramMarkdown = WshNamed('Markdown');
    if (WshNamed.Exists('Markdown') && paramMarkdown && paramMarkdown.length) {
      return {
        Markdown: paramMarkdown
      }
    }
    var param = { Set: WshNamed.Exists('Set') };
    if (param.Set) {
      var noIconParam = WshNamed('Set');
      var isNoIconParam = false;
      param.NoIcon = noIconParam && (isNoIconParam = /^NoIcon$/i.test(noIconParam));
      if (noIconParam == undefined || isNoIconParam) {
        return param;
      }
    }
    param = { Unset: WshNamed.Exists('Unset') };
    if (param.Unset && WshNamed('Unset') == undefined) {
      return param;
    }
  } else if (paramCount == 0) {
    return {
      Set: true,
      NoIcon: false
    }
  }
  var helpText = '';
  helpText += 'The MarkdownToHtml shortcut launcher.\n';
  helpText += 'It starts the shortcut menu target script in a hidden window.\n\n';
  helpText += 'Syntax:\n';
  helpText += '  Convert-MarkdownToHtml.js /Markdown:<markdown file path>\n';
  helpText += '  Convert-MarkdownToHtml.js [/Set[:NoIcon]]\n';
  helpText += '  Convert-MarkdownToHtml.js /Unset\n';
  helpText += '  Convert-MarkdownToHtml.js /Help\n\n';
  helpText += "<markdown file path>  The selected markdown's file path.\n";
  helpText += '                 Set  Configure the shortcut menu in the registry.\n';
  helpText += '              NoIcon  Specifies that the icon is not configured.\n';
  helpText += '               Unset  Removes the shortcut menu.\n';
  helpText += '                Help  Show the help doc.\n';
  popup(helpText);
  quit(1);
}

// #endregion