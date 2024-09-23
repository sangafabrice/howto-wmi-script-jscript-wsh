/**
 * @file Launches the shortcut target PowerShell script with the selected markdown as an argument.
 * @version 0.0.1.0
 */

// #region: header of utils.js
// Constants and variables.

/** @constant */
var WINDOW_STYLE_HIDDEN = 0;
/** @constant */
var BUTTONS_OKONLY = 0;
/** @constant */
var POPUP_NORMAL = 0;

/** @typedef */
var FileSystemObject = new ActiveXObject('Scripting.FileSystemObject');
/** @typedef */
var WshShell = new ActiveXObject('WScript.Shell');
/** @typedef */
var Shell32 = new ActiveXObject('Shell.Application');
/** @typedef */
var StdRegProv = GetObject('winmgmts:StdRegProv');

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
  // #region: process.js

  // #region: Process type definition

  var Process = (function() {
    /** @constructor */
    function Process() { }

    /**
     * Start a specified program with its arguments.
     * @param {ProcessStartup} startInfo the process startup information.
     * @returns {number} the process creation status.
     */
    Process.Start = function (startInfo) {
      /** @typedef */
      var Win32_Process = GetObject('winmgmts:Win32_Process');
      startInfo.StartupInfo.ShowWindow = startInfo.WindowStyle;
      try {
        return Win32_Process.Create(startInfo.CommandLine, null, startInfo.StartupInfo);
      } finally {
        Win32_Process = null;
        startInfo = null;
      }
    }

    return Process;
  })();

  // #endregion

  // #region: ProcessStartup type definition

  var ProcessStartup = (function() {
    /** @typedef */
    var Win32_ProcessStartup = null;

    /**
     * @constructor
     * @param {string} commandLine the command line to execute.
     */
    function ProcessStartup(commandLine) {
      Win32_ProcessStartup = GetObject('winmgmts:Win32_ProcessStartup');
      /** The Win32_ProcessStartup instance. */
      this.StartupInfo = Win32_ProcessStartup.SpawnInstance_();
      /** The window style of the starting process. */
      this.WindowStyle = this.StartupInfo.ShowWindow;
      /** The command line to execute. */
      this.CommandLine = commandLine;
    }

    ProcessStartup.prototype.Dispose = function() {
      this.StartupInfo = null;
      Win32_ProcessStartup = null;
    }

    return ProcessStartup;
  })();

  // #endregion

  // #endregion

  /** @constant */
  var CMD_LINE_FORMAT = 'C:\\Windows\\System32\\cmd.exe /d /c ""{0}" "{1}""';
  var startInfo = new ProcessStartup(format(CMD_LINE_FORMAT, Package.IconLink.Path, Param.Markdown));
  startInfo.WindowStyle = WINDOW_STYLE_HIDDEN;
  Process.Start(startInfo);
  startInfo.Dispose();
  quit(0);
}

/** Configuration and settings. */
if (Param.Set ^ Param.Unset) {
  // #region: setup.js
  // Methods for managing the shortcut menu option: install and uninstall.

  /** @typedef */
  var Setup = (function() {
    var HKCU = 0x80000001;
    var VERB_KEY = 'SOFTWARE\\Classes\\SystemFileAssociations\\.md\\shell\\cthtml';
    var ICON_VALUENAME = 'Icon';

    return {
      /** Configure the shortcut menu in the registry. */
      Set: function () {
        var COMMAND_KEY = VERB_KEY + '\\command';
        var command = format('{0} "{1}" /Markdown:"%1"', WSH.FullName.replace(/\\cscript\.exe$/i, '\\wscript.exe'), WSH.ScriptFullName);
        StdRegProv.CreateKey(HKCU, COMMAND_KEY);
        StdRegProv.SetStringValue(HKCU, COMMAND_KEY, null, command);
        StdRegProv.SetStringValue(HKCU, VERB_KEY, null, 'Convert to &HTML');
      },

      /**
       * Add an icon to the shortcut menu in the registry.
       * @param {string} menuIconPath is the shortcut menu icon file path.
       */
      AddIcon: function (menuIconPath) {
        StdRegProv.SetStringValue(HKCU, VERB_KEY, ICON_VALUENAME, menuIconPath);
      },

      /** Remove the shortcut icon menu. */
      RemoveIcon: function () {
        StdRegProv.DeleteValue(HKCU, VERB_KEY, ICON_VALUENAME);
      },

      /** Remove the shortcut menu by removing the verb key and subkeys. */
      Unset: function () {
        var stdRegProvMethods = StdRegProv.Methods_;
        var enumKeyMethod = stdRegProvMethods('EnumKey');
        var enumKeyMethodParams = enumKeyMethod.InParameters;
        var inParams = enumKeyMethodParams.SpawnInstance_();
        inParams.hDefKey = HKCU;
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
          StdRegProv.DeleteKey(HKCU, key);
        })(VERB_KEY);
        inParams = null;
        enumKeyMethodParams = null;
        enumKeyMethod = null;
        stdRegProvMethods = null;
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
  Shell32.ShellExecute(Package.MessageBoxLinkPath, format('""""{0}"""" {1} {2}', messageText.replace(/"/g, "'"), popupButtons, popupType), null, 'open', WINDOW_STYLE_HIDDEN);
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
  StdRegProv = null;
  Shell32 = null;
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
  var POWERSHELL_SUBKEY = 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\pwsh.exe';
  /** The project resources directory path. */
  var resourcePath = FileSystemObject.BuildPath(ScriptRoot, 'rsc');

  /** The powershell core runtime path. */
  var pwshExePath = (function() {
    var stdRegProvMethods = StdRegProv.Methods_;
    var getStringValueMethod = stdRegProvMethods('GetStringValue');
    var getStringValueMethodParams = getStringValueMethod.InParameters;
    var inParams = getStringValueMethodParams.SpawnInstance_();
    inParams.sSubKeyName = POWERSHELL_SUBKEY;
    var outParams = StdRegProv.ExecMethod_(getStringValueMethod.Name, inParams);
    try {
      return outParams.sValue;
    } finally {
      outParams = null;
      inParams = null;
      getStringValueMethodParams = null;
      getStringValueMethod = null;
      stdRegProvMethods = null;
    }
  })();
  /** The shortcut target powershell script path. */
  var pwshScriptPath = FileSystemObject.BuildPath(resourcePath, 'cvmd2html.ps1');
  var msgBoxLinkPath = FileSystemObject.BuildPath(ScriptRoot, 'MsgBox.lnk');
  var menuIconPath = FileSystemObject.BuildPath(resourcePath, 'menu.ico');

  /** The icon link parent directory name. */
  var iconLinkDirName = WshShell.SpecialFolders('StartMenu');
  /** The icon link name. */
  var iconLinkName = 'cvmd2html.lnk';

  return {
    /** The shortcut menu icon path. */
    MenuIconPath: menuIconPath,
    /** The message box link path. */
    MessageBoxLinkPath: msgBoxLinkPath,

    /** Represents an adapted link object. */
    IconLink: {
      /** The icon link path. */
      Path: FileSystemObject.BuildPath(iconLinkDirName, iconLinkName),

      /** Create the custom icon link file. */
      Create: function () {
        var txtFile = FileSystemObject.CreateTextFile(this.Path);
        txtFile.Close();
        txtFile = null;
        var folderItem = Shell32.NameSpace(iconLinkDirName);
        var fileItem = folderItem.ParseName(iconLinkName);
        var link = fileItem.GetLink;
        link.Path = pwshExePath;
        link.Arguments = format('-ep Bypass -nop -w Hidden -f "{0}" -Markdown', pwshScriptPath);
        link.SetIconLocation(menuIconPath, 0);
        link.Save();
        link = null;
        fileItem = null;
        folderItem = null;
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
        var linkItem;
        /** @typedef */
        var Win32_ShortcutFile = GetObject('winmgmts:Win32_ShortcutFile');
        var allShorcutFiles = Win32_ShortcutFile.Instances_();
        var linkEnumerator = new Enumerator(allShorcutFiles);
        var minPwshExePath = pwshExePath.toLowerCase();
        var minIconLinkPath = this.Path.toLowerCase();
        try {
          while (!linkEnumerator.atEnd()) {
            if ((linkItem = linkEnumerator.item()).Name.toLowerCase() == minIconLinkPath) {
              return linkItem.Target.toLowerCase() == minPwshExePath;
            }
            linkEnumerator.moveNext()
          }
          return false;
        } finally {
          linkItem = null;
          linkEnumerator = null;
          allShorcutFiles = null;
          Win32_ShortcutFile = null;
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