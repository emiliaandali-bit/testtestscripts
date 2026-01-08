#target photoshop

(function () {
    if (parseInt(app.version, 10) < 23) {
        alert("This script targets Photoshop 2022+ (v23). Some features may not work in older versions.");
    }

    var SCRIPT_NAME = "Bulk Mockup Builder";
    var DESIGN_LAYER_NAME = "Design";
    var logFile = null;
    var originalRulerUnits = app.preferences.rulerUnits;
    var originalTypeUnits = app.preferences.typeUnits;
    var originalDisplayDialogs = app.displayDialogs;

    function setDialogModesNo() {
        app.displayDialogs = DialogModes.NO;
    }

    function restorePrefs() {
        app.preferences.rulerUnits = originalRulerUnits;
        app.preferences.typeUnits = originalTypeUnits;
        app.displayDialogs = originalDisplayDialogs;
    }

    function timestamp() {
        var d = new Date();
        function pad(n) { return n < 10 ? "0" + n : n; }
        return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + " " + pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds());
    }

    function initLog(exportFolder) {
        try {
            logFile = new File(exportFolder.fsName + "/bulk-mockup-log.txt");
            logFile.open("a");
            logFile.writeln("\n==== " + SCRIPT_NAME + " Run: " + timestamp() + " ====");
            logFile.close();
        } catch (e) {
            // ignore logging errors
        }
    }

    function logLine(msg) {
        try {
            if (!logFile) { return; }
            logFile.open("a");
            logFile.writeln("[" + timestamp() + "] " + msg);
            logFile.close();
        } catch (e) {
            // ignore logging errors
        }
    }

    function pickFolder(title) {
        return Folder.selectDialog(title);
    }

    function getFilesByExtensions(folder, extensions) {
        var files = folder.getFiles(function (f) {
            if (f instanceof File) {
                var ext = f.name.toLowerCase().split('.').pop();
                for (var i = 0; i < extensions.length; i++) {
                    if (ext === extensions[i]) return true;
                }
            }
            return false;
        });
        return files;
    }

    function findLayerByName(layerSet, name) {
        for (var i = 0; i < layerSet.layers.length; i++) {
            var layer = layerSet.layers[i];
            if (layer.typename === "ArtLayer" && layer.name === name) {
                return layer;
            }
            if (layer.typename === "LayerSet") {
                var found = findLayerByName(layer, name);
                if (found) return found;
            }
        }
        return null;
    }

    function isSmartObjectLayer(layer) {
        try {
            return layer.kind === LayerKind.SMARTOBJECT;
        } catch (e) {
            return false;
        }
    }

    function selectLayer(layer) {
        var ref = new ActionReference();
        ref.putIdentifier(charIDToTypeID("Lyr "), layer.id);
        var desc = new ActionDescriptor();
        desc.putReference(charIDToTypeID("null"), ref);
        executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
    }

    function replaceSmartObjectContents(layer, file) {
        selectLayer(layer);
        var idplacedLayerReplaceContents = stringIDToTypeID("placedLayerReplaceContents");
        var desc = new ActionDescriptor();
        desc.putPath(charIDToTypeID("null"), file);
        executeAction(idplacedLayerReplaceContents, desc, DialogModes.NO);
    }

    function editSmartObjectContents() {
        var idplacedLayerEditContents = stringIDToTypeID("placedLayerEditContents");
        executeAction(idplacedLayerEditContents, new ActionDescriptor(), DialogModes.NO);
    }

    function fitOrFillActiveLayer(doc, mode) {
        var layer = doc.activeLayer;
        var b = layer.bounds; // UnitValues
        var l = b[0].as("px"), t = b[1].as("px"), r = b[2].as("px"), btm = b[3].as("px");
        var layerW = r - l;
        var layerH = btm - t;
        var docW = doc.width.as("px");
        var docH = doc.height.as("px");
        if (layerW === 0 || layerH === 0) { return; }
        var scaleX = docW / layerW * 100;
        var scaleY = docH / layerH * 100;
        var scale = (mode === "fit") ? Math.min(scaleX, scaleY) : Math.max(scaleX, scaleY);
        layer.resize(scale, scale, AnchorPosition.MIDDLECENTER);
        // center
        var nb = layer.bounds;
        var nl = nb[0].as("px"), nt = nb[1].as("px"), nr = nb[2].as("px"), nbm = nb[3].as("px");
        var nx = (nl + nr) / 2;
        var ny = (nt + nbm) / 2;
        var dx = docW / 2 - nx;
        var dy = docH / 2 - ny;
        layer.translate(dx, dy);
    }

    function editSOAndFit(mode) {
        editSmartObjectContents();
        var soDoc = app.activeDocument;
        // Assume top layer is replaced content
        soDoc.activeLayer = soDoc.layers[0];
        fitOrFillActiveLayer(soDoc, mode);
        soDoc.save();
        soDoc.close(SaveOptions.SAVECHANGES);
    }

    function sanitizeFilename(name, sep) {
        var illegal = /[\\\/:*?"<>|]/g;
        var cleaned = name.replace(illegal, "").replace(/\s+/g, " ").replace(/\s*$/g, "").replace(/^\s*/g, "");
        var escSep = sep.replace(/[-/\\^$*+?.()|[\]{}]/g, "\\$&");
        var multiSep = new RegExp(escSep + "+", "g");
        cleaned = cleaned.replace(multiSep, sep);
        return cleaned;
    }

    function buildFilename(opts, mockupName, designName) {
        var parts = [];
        var appendName = opts.appendDesign ? designName : "";
        if (opts.appendDesign && opts.appendPosition === "prefix") {
            parts.push(appendName);
        }
        if (opts.prefix) { parts.push(opts.prefix); }
        if (opts.includeMockup) { parts.push(mockupName); }
        if (opts.appendDesign && opts.appendPosition === "suffix") {
            parts.push(appendName);
        }
        if (opts.suffix) { parts.push(opts.suffix); }
        var sep = opts.separator || "_";
        var name = parts.join(sep);
        name = sanitizeFilename(name, sep);
        return name;
    }

    function ensureFolder(folder) {
        if (!folder.exists) {
            folder.create();
        }
    }

    function incrementFilenameIfNeeded(folder, baseName, ext) {
        var i = 1;
        var candidate = new File(folder.fsName + "/" + baseName + ext);
        while (candidate.exists) {
            candidate = new File(folder.fsName + "/" + baseName + "_" + i + ext);
            i++;
        }
        return candidate;
    }

    function exportPNG(doc, file, opts) {
        var exp = new ExportOptionsSaveForWeb();
        exp.format = SaveDocumentType.PNG;
        exp.PNG8 = false;
        exp.transparency = opts.pngTransparent;
        exp.includeProfile = opts.pngEmbedProfile;
        doc.exportDocument(file, ExportType.SAVEFORWEB, exp);
    }

    function exportJPEG(doc, file, opts) {
        var exp = new ExportOptionsSaveForWeb();
        exp.format = SaveDocumentType.JPEG;
        exp.quality = opts.jpegQuality;
        exp.includeProfile = true;
        exp.optimized = true;
        exp.progressive = false;
        doc.exportDocument(file, ExportType.SAVEFORWEB, exp);
    }

    function exportPSD(doc, file, opts) {
        var psdOpts = new PhotoshopSaveOptions();
        psdOpts.layers = true;
        psdOpts.maximizeCompatibility = opts.psdMaxCompat;
        doc.saveAs(file, psdOpts, true, Extension.LOWERCASE);
    }

    function exportWebP(doc, file, opts) {
        try {
            var desc = new ActionDescriptor();
            var desc2 = new ActionDescriptor();
            desc2.putBoolean(stringIDToTypeID("lossless"), opts.webpLossless);
            if (!opts.webpLossless) {
                desc2.putInteger(stringIDToTypeID("quality"), opts.webpQuality);
            }
            desc.putObject(stringIDToTypeID("as"), stringIDToTypeID("webpFormat"), desc2);
            desc.putPath(stringIDToTypeID("in"), file);
            executeAction(stringIDToTypeID("save"), desc, DialogModes.NO);
            return true;
        } catch (e) {
            return false;
        }
    }

    function buildOutputFolder(exportRoot, opts, mockupName, designName) {
        var folder = exportRoot;
        if (opts.subfolderPerDesign) {
            folder = new Folder(folder.fsName + "/" + designName);
            ensureFolder(folder);
        }
        if (opts.subfolderPerMockup) {
            folder = new Folder(folder.fsName + "/" + mockupName);
            ensureFolder(folder);
        }
        return folder;
    }

    function updatePreview(previewText, opts, mockupName, designName, ext) {
        var name = buildFilename(opts, mockupName, designName);
        previewText.text = "Preview: " + name + ext;
    }

    function createUI() {
        var win = new Window("dialog", SCRIPT_NAME);
        win.alignChildren = "fill";

        var foldersPanel = win.add("panel", undefined, "Folders");
        foldersPanel.orientation = "column";
        foldersPanel.alignChildren = "fill";

        function folderRow(label) {
            var group = foldersPanel.add("group");
            group.add("statictext", undefined, label).preferredSize.width = 140;
            var field = group.add("edittext", undefined, "");
            field.characters = 40;
            var btn = group.add("button", undefined, "Browse");
            return { group: group, field: field, btn: btn };
        }

        var mockupRow = folderRow("Mockups Folder:");
        var designRow = folderRow("Designs Folder:");
        var exportRow = folderRow("Export Folder:");

        var exportPanel = win.add("panel", undefined, "Export Settings");
        exportPanel.orientation = "column";
        exportPanel.alignChildren = "left";

        var formatGroup = exportPanel.add("group");
        formatGroup.add("statictext", undefined, "Format:");
        var formatDropdown = formatGroup.add("dropdownlist", undefined, ["PNG", "JPEG", "WEBP", "PSD"]);
        formatDropdown.selection = 0;

        var pngGroup = exportPanel.add("group");
        pngGroup.add("statictext", undefined, "PNG Options:");
        var pngTransparent = pngGroup.add("checkbox", undefined, "Transparent");
        pngTransparent.value = true;
        var pngEmbedProfile = pngGroup.add("checkbox", undefined, "Embed Profile");
        pngEmbedProfile.value = true;

        var jpegGroup = exportPanel.add("group");
        jpegGroup.add("statictext", undefined, "JPEG Quality:");
        var jpegSlider = jpegGroup.add("slider", undefined, 10, 1, 12);
        jpegSlider.preferredSize.width = 120;
        var jpegValue = jpegGroup.add("statictext", undefined, "10");

        var webpGroup = exportPanel.add("group");
        webpGroup.add("statictext", undefined, "WebP:");
        var webpMode = webpGroup.add("dropdownlist", undefined, ["Lossy", "Lossless"]);
        webpMode.selection = 0;
        var webpQualityGroup = exportPanel.add("group");
        webpQualityGroup.add("statictext", undefined, "WebP Quality:");
        var webpQuality = webpQualityGroup.add("slider", undefined, 80, 0, 100);
        webpQuality.preferredSize.width = 120;
        var webpValue = webpQualityGroup.add("statictext", undefined, "80");

        var psdGroup = exportPanel.add("group");
        psdGroup.add("statictext", undefined, "PSD Options:");
        var psdMaxCompat = psdGroup.add("checkbox", undefined, "Maximize Compatibility");
        psdMaxCompat.value = true;

        var fitGroup = exportPanel.add("group");
        fitGroup.add("statictext", undefined, "Fit/Fill:");
        var fitDropdown = fitGroup.add("dropdownlist", undefined, ["Fill", "Fit"]);
        fitDropdown.selection = 0;

        var namingPanel = win.add("panel", undefined, "Naming");
        namingPanel.orientation = "column";
        namingPanel.alignChildren = "fill";

        var prefixGroup = namingPanel.add("group");
        prefixGroup.add("statictext", undefined, "Prefix:");
        var prefixInput = prefixGroup.add("edittext", undefined, "");
        prefixInput.characters = 20;

        var suffixGroup = namingPanel.add("group");
        suffixGroup.add("statictext", undefined, "Suffix:");
        var suffixInput = suffixGroup.add("edittext", undefined, "");
        suffixInput.characters = 20;

        var appendGroup = namingPanel.add("group");
        var appendDesign = appendGroup.add("checkbox", undefined, "Append original design filename");
        appendDesign.value = true;
        var appendPos = appendGroup.add("dropdownlist", undefined, ["as suffix", "as prefix"]);
        appendPos.selection = 0;

        var includeMockupGroup = namingPanel.add("group");
        var includeMockup = includeMockupGroup.add("checkbox", undefined, "Include mockup PSD name in output filename");
        includeMockup.value = true;

        var separatorGroup = namingPanel.add("group");
        separatorGroup.add("statictext", undefined, "Separator:");
        var separatorInput = separatorGroup.add("edittext", undefined, "_");
        separatorInput.characters = 5;

        var previewGroup = namingPanel.add("group");
        var previewText = previewGroup.add("statictext", undefined, "Preview: ");

        var subfolderGroup = namingPanel.add("group");
        var subPerDesign = subfolderGroup.add("checkbox", undefined, "Create subfolder per design in export folder");
        var subPerMockup = subfolderGroup.add("checkbox", undefined, "Create subfolder per mockup");

        var overwriteGroup = namingPanel.add("group");
        var overwriteCheckbox = overwriteGroup.add("checkbox", undefined, "Overwrite existing files");
        overwriteCheckbox.value = false;

        var dryRunGroup = namingPanel.add("group");
        var dryRunCheckbox = dryRunGroup.add("checkbox", undefined, "Dry run (log only)");
        dryRunCheckbox.value = false;

        var progressPanel = win.add("panel", undefined, "Progress");
        progressPanel.orientation = "column";
        progressPanel.alignChildren = "fill";
        var progressBar = progressPanel.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize.width = 300;
        var statusText = progressPanel.add("statictext", undefined, "Idle");

        var buttonsGroup = win.add("group");
        buttonsGroup.alignment = "right";
        var runBtn = buttonsGroup.add("button", undefined, "Run");
        var cancelBtn = buttonsGroup.add("button", undefined, "Cancel", { name: "cancel" });

        function updateOptionsVisibility() {
            var fmt = formatDropdown.selection.text;
            pngGroup.visible = fmt === "PNG";
            jpegGroup.visible = fmt === "JPEG";
            webpGroup.visible = fmt === "WEBP";
            webpQualityGroup.visible = fmt === "WEBP" && webpMode.selection.text === "Lossy";
            psdGroup.visible = fmt === "PSD";
            win.layout.layout(true);
        }

        jpegSlider.onChanging = function () { jpegValue.text = Math.round(jpegSlider.value).toString(); };
        webpQuality.onChanging = function () { webpValue.text = Math.round(webpQuality.value).toString(); };
        formatDropdown.onChange = updateOptionsVisibility;
        webpMode.onChange = updateOptionsVisibility;

        function updatePreviewLine() {
            var opts = getNamingOptions();
            updatePreview(previewText, opts, "Mockup", "Design", "." + formatDropdown.selection.text.toLowerCase());
        }

        function getNamingOptions() {
            return {
                prefix: prefixInput.text,
                suffix: suffixInput.text,
                appendDesign: appendDesign.value,
                appendPosition: appendPos.selection.text === "as prefix" ? "prefix" : "suffix",
                includeMockup: includeMockup.value,
                separator: separatorInput.text || "_",
                subfolderPerDesign: subPerDesign.value,
                subfolderPerMockup: subPerMockup.value,
                overwrite: overwriteCheckbox.value
            };
        }

        var inputsToWatch = [prefixInput, suffixInput, appendDesign, appendPos, includeMockup, separatorInput, formatDropdown];
        for (var i = 0; i < inputsToWatch.length; i++) {
            inputsToWatch[i].onChange = updatePreviewLine;
        }
        updateOptionsVisibility();
        updatePreviewLine();

        mockupRow.btn.onClick = function () {
            var f = pickFolder("Select Mockups Folder");
            if (f) mockupRow.field.text = f.fsName;
        };
        designRow.btn.onClick = function () {
            var f = pickFolder("Select Designs Folder");
            if (f) designRow.field.text = f.fsName;
        };
        exportRow.btn.onClick = function () {
            var f = pickFolder("Select Export Folder");
            if (f) exportRow.field.text = f.fsName;
        };

        runBtn.onClick = function () { win.close(1); };
        cancelBtn.onClick = function () { win.close(0); };

        win.center();
        var result = win.show();
        if (result !== 1) { return null; }

        return {
            mockupsFolder: new Folder(mockupRow.field.text),
            designsFolder: new Folder(designRow.field.text),
            exportFolder: new Folder(exportRow.field.text),
            format: formatDropdown.selection.text,
            jpegQuality: Math.round(jpegSlider.value),
            webpLossless: webpMode.selection.text === "Lossless",
            webpQuality: Math.round(webpQuality.value),
            pngTransparent: pngTransparent.value,
            pngEmbedProfile: pngEmbedProfile.value,
            psdMaxCompat: psdMaxCompat.value,
            fitMode: fitDropdown.selection.text.toLowerCase(),
            naming: getNamingOptions(),
            dryRun: dryRunCheckbox.value,
            progressBar: progressBar,
            statusText: statusText
        };
    }

    function runBatch(opts) {
        if (!opts.mockupsFolder.exists || !opts.designsFolder.exists || !opts.exportFolder.exists) {
            alert("Please select valid folders for mockups, designs, and export.");
            return;
        }

        initLog(opts.exportFolder);
        setDialogModesNo();
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.PIXELS;

        var mockupFiles = getFilesByExtensions(opts.mockupsFolder, ["psd"]);
        var designFiles = getFilesByExtensions(opts.designsFolder, ["png", "jpg", "jpeg", "webp", "tif", "tiff", "psd"]);

        var total = mockupFiles.length * designFiles.length;
        if (total === 0) {
            alert("No mockup PSDs or design files found.");
            return;
        }

        var count = 0;
        for (var i = 0; i < mockupFiles.length; i++) {
            var mockupFile = mockupFiles[i];
            var mockupName = mockupFile.name.replace(/\.[^\.]+$/, "");
            for (var j = 0; j < designFiles.length; j++) {
                var designFile = designFiles[j];
                var designName = designFile.name.replace(/\.[^\.]+$/, "");
                count++;
                var status = "Processing: " + mockupFile.name + " + " + designFile.name;
                opts.statusText.text = status;
                opts.progressBar.value = Math.round(count / total * 100);

                try {
                    if (opts.dryRun) {
                        logLine("DRY RUN: " + status);
                        continue;
                    }

                    var doc = app.open(mockupFile);
                    var designLayer = findLayerByName(doc, DESIGN_LAYER_NAME);
                    if (!designLayer || !isSmartObjectLayer(designLayer)) {
                        logLine("ERROR: Smart Object layer named '" + DESIGN_LAYER_NAME + "' not found in " + mockupFile.name);
                        doc.close(SaveOptions.DONOTSAVECHANGES);
                        continue;
                    }

                    replaceSmartObjectContents(designLayer, designFile);
                    editSOAndFit(opts.fitMode);

                    var outputFolder = buildOutputFolder(opts.exportFolder, opts.naming, mockupName, designName);
                    var baseName = buildFilename(opts.naming, mockupName, designName);
                    var ext = "." + opts.format.toLowerCase();
                    var outputFile = new File(outputFolder.fsName + "/" + baseName + ext);
                    if (!opts.naming.overwrite && outputFile.exists) {
                        outputFile = incrementFilenameIfNeeded(outputFolder, baseName, ext);
                    }

                    if (opts.format === "PNG") {
                        exportPNG(doc, outputFile, opts);
                    } else if (opts.format === "JPEG") {
                        exportJPEG(doc, outputFile, opts);
                    } else if (opts.format === "PSD") {
                        exportPSD(doc, outputFile, opts);
                    } else if (opts.format === "WEBP") {
                        var ok = exportWebP(doc, outputFile, opts);
                        if (!ok) {
                            logLine("WARNING: WebP export not available, falling back to PNG for " + baseName);
                            var fallbackFile = new File(outputFolder.fsName + "/" + baseName + ".png");
                            exportPNG(doc, fallbackFile, opts);
                        }
                    }

                    doc.close(SaveOptions.DONOTSAVECHANGES);
                    logLine("SUCCESS: " + outputFile.fsName);
                } catch (e) {
                    logLine("ERROR: " + status + " | " + e.toString());
                    try {
                        if (app.documents.length > 0) {
                            app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
                        }
                    } catch (closeErr) {}
                }
            }
        }

        opts.statusText.text = "Done";
        opts.progressBar.value = 100;
    }

    var options = createUI();
    if (options) {
        try {
            runBatch(options);
        } finally {
            restorePrefs();
        }
    }
})();
