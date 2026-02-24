"""
MultiMeshExport – Fusion 360 Add-in
Export multiple bodies as STL files in a single operation.

Features
--------
* Toggle-select any body in the design (or use "Select All").
* Choose High / Medium / Low mesh quality (default: High).
* Pick a save folder (default: Downloads).
* Automatically overwrites existing files with the same name.
* Progress bar with cancel support.
"""

import adsk.core
import adsk.fusion
import traceback
import os
import json
from pathlib import Path


# ──────────────────────────────────────────────────────────────────────────────
# Globals
# ──────────────────────────────────────────────────────────────────────────────
_app: adsk.core.Application = None
_ui: adsk.core.UserInterface = None
_handlers = []                  # prevent GC of event handlers
_updating = False               # guard against recursive input-changed events
_custom_names = {}              # body entityToken → user-chosen save name

CMD_ID   = 'multiMeshExportCmd'
CMD_NAME = 'Multi Mesh Export'
CMD_DESC = 'Select and export multiple bodies as STL files at once.'

QUALITY = {
    'High':   adsk.fusion.MeshRefinementSettings.MeshRefinementHigh,
    'Medium': adsk.fusion.MeshRefinementSettings.MeshRefinementMedium,
    'Low':    adsk.fusion.MeshRefinementSettings.MeshRefinementLow,
}

_SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              'settings.json')


def _load_settings():
    """Load persisted settings from disk (returns dict)."""
    try:
        with open(_SETTINGS_FILE, 'r') as f:
            return json.load(f)
    except Exception:
        return {}


def _save_settings(data):
    """Merge *data* into the persisted settings file."""
    settings = _load_settings()
    settings.update(data)
    try:
        with open(_SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def _downloads_folder():
    """Return the current user's Downloads folder."""
    return str(Path.home() / 'Downloads')


def _all_bodies(design):
    """Collect every BRepBody from every component in the design."""
    bodies = []
    for comp in design.allComponents:
        for body in comp.bRepBodies:
            bodies.append(body)
    return bodies


def _safe_filename(name):
    """Strip characters that are illegal in Windows / macOS file names."""
    return ''.join(c for c in name if c not in r'\/:*?"<>|').strip() or 'body'


def _rebuild_name_list(inputs):
    """Rebuild the editable save-name list to match the current body selection."""
    sel = adsk.core.SelectionCommandInput.cast(inputs.itemById('bodySelection'))
    group = adsk.core.GroupCommandInput.cast(inputs.itemById('bodyListGroup'))
    if not sel or not group:
        return

    children = group.children

    # Remove old rows
    while children.count > 0:
        children.item(children.count - 1).deleteMe()

    # Create one editable name field per selected body
    for i in range(sel.selectionCount):
        body = sel.selection(i).entity
        token = body.entityToken
        default = _custom_names.get(token, _safe_filename(body.name))
        children.addStringValueInput(
            'saveName_{}'.format(i), body.name, default)


# ──────────────────────────────────────────────────────────────────────────────
# Event Handlers
# ──────────────────────────────────────────────────────────────────────────────
class _OnCommandCreated(adsk.core.CommandCreatedEventHandler):
    """Build the command dialog UI."""

    def notify(self, args):
        try:
            cmd    = args.command
            cmd.isRepeatable = False
            inputs = cmd.commandInputs

            design = adsk.fusion.Design.cast(_app.activeProduct)
            if not design:
                _ui.messageBox(
                    'No active design.\nOpen or create a design first.',
                    CMD_NAME,
                    adsk.core.MessageBoxButtonTypes.OKButtonType,
                    adsk.core.MessageBoxIconTypes.InformationIconType,
                )
                return

            # ── Body selection (multi-select, bodies only) ──
            sel = inputs.addSelectionInput(
                'bodySelection', 'Bodies',
                'Select one or more bodies to export as STL')
            sel.addSelectionFilter('Bodies')
            sel.setSelectionLimits(0, 0)        # unlimited

            # ── Select / Deselect all ──
            inputs.addBoolValueInput('selectAll', 'Select All Bodies',
                                     True, '', False)

            # ── Editable save names (populated when bodies are selected) ──
            _custom_names.clear()
            group = inputs.addGroupCommandInput('bodyListGroup',
                                                'Export Names')
            group.isExpanded = True

            # ── Mesh quality ──
            dd = inputs.addDropDownCommandInput(
                'quality', 'Mesh Quality',
                adsk.core.DropDownStyles.TextListDropDownStyle)
            dd.listItems.add('High',   True,  '')
            dd.listItems.add('Medium', False, '')
            dd.listItems.add('Low',    False, '')

            # ── Save location ──
            saved_path = _load_settings().get('outputPath', _downloads_folder())
            inputs.addStringValueInput('outputPath', 'Save Location',
                                       saved_path)
            browse_btn = inputs.addBoolValueInput('browseFolder', '📁 Browse ', False, '', False)
            browse_btn.isFullWidth = True

            # ── Wire up handlers ──
            h1 = _OnExecute();        cmd.execute.add(h1);        _handlers.append(h1)
            h2 = _OnInputChanged();   cmd.inputChanged.add(h2);   _handlers.append(h2)
            h3 = _OnValidateInputs(); cmd.validateInputs.add(h3); _handlers.append(h3)

        except Exception:
            _ui.messageBox(traceback.format_exc())


class _OnInputChanged(adsk.core.InputChangedEventHandler):
    """React to UI changes: Select All, Browse, and sync state."""

    def notify(self, args):
        global _updating
        if _updating:
            return
        _updating = True
        try:
            changed = args.input
            inputs  = args.inputs

            # ── Select All toggle ──
            if changed.id == 'selectAll':
                sel = inputs.itemById('bodySelection')
                chk = adsk.core.BoolValueCommandInput.cast(changed)
                design = adsk.fusion.Design.cast(_app.activeProduct)
                if not design:
                    return
                if chk.value:
                    for body in _all_bodies(design):
                        try:
                            sel.addSelection(body)
                        except Exception:
                            pass        # suppressed / invisible bodies
                else:
                    sel.clearSelection()
                _rebuild_name_list(inputs)

            # ── Keep "Select All" in sync with manual selection ──
            elif changed.id == 'bodySelection':
                sel = adsk.core.SelectionCommandInput.cast(changed)
                chk = inputs.itemById('selectAll')
                design = adsk.fusion.Design.cast(_app.activeProduct)
                if design:
                    total = len(_all_bodies(design))
                    chk.value = (total > 0 and sel.selectionCount == total)
                _rebuild_name_list(inputs)

            # ── Save custom name when the user edits a row ──
            elif changed.id.startswith('saveName_'):
                idx = int(changed.id.split('_')[1])
                sel = adsk.core.SelectionCommandInput.cast(
                    inputs.itemById('bodySelection'))
                if sel and idx < sel.selectionCount:
                    body = sel.selection(idx).entity
                    _custom_names[body.entityToken] = (
                        adsk.core.StringValueCommandInput.cast(changed).value)

            # ── Browse for folder ──
            elif changed.id == 'browseFolder':
                chk = adsk.core.BoolValueCommandInput.cast(changed)
                if chk.value:
                    dlg = _ui.createFolderDialog()
                    dlg.title = 'Choose Export Folder'
                    path_inp = adsk.core.StringValueCommandInput.cast(
                        inputs.itemById('outputPath'))
                    if os.path.isdir(path_inp.value):
                        dlg.initialDirectory = path_inp.value
                    if dlg.showDialog() == adsk.core.DialogResults.DialogOK:
                        path_inp.value = dlg.folder
                    chk.value = False       # reset so it can be clicked again

        except Exception:
            _ui.messageBox(traceback.format_exc())
        finally:
            _updating = False


class _OnValidateInputs(adsk.core.ValidateInputsEventHandler):
    """OK button is enabled only when ≥1 body is selected and a path is set."""

    def notify(self, args):
        try:
            cmd    = adsk.core.Command.cast(args.firingEvent.sender)
            inputs = cmd.commandInputs
            sel    = adsk.core.SelectionCommandInput.cast(
                         inputs.itemById('bodySelection'))
            path   = adsk.core.StringValueCommandInput.cast(
                         inputs.itemById('outputPath'))
            args.areInputsValid = (
                sel.selectionCount > 0
                and len(path.value.strip()) > 0
            )
        except Exception:
            args.areInputsValid = False


class _OnExecute(adsk.core.CommandEventHandler):
    """Perform the actual STL export for every selected body."""

    def notify(self, args):
        try:
            inputs   = args.command.commandInputs
            sel      = adsk.core.SelectionCommandInput.cast(
                           inputs.itemById('bodySelection'))
            q_dd     = adsk.core.DropDownCommandInput.cast(
                           inputs.itemById('quality'))
            path_inp = adsk.core.StringValueCommandInput.cast(
                           inputs.itemById('outputPath'))

            bodies     = [sel.selection(i).entity
                          for i in range(sel.selectionCount)]
            quality    = QUALITY.get(q_dd.selectedItem.name, QUALITY['High'])
            output_dir = path_inp.value.strip()

            if not bodies:
                _ui.messageBox('No bodies selected.', CMD_NAME)
                return

            os.makedirs(output_dir, exist_ok=True)
            _save_settings({'outputPath': output_dir})

            design     = adsk.fusion.Design.cast(_app.activeProduct)
            export_mgr = design.exportManager

            # ── Build file names from editable save-name inputs ──────────
            group = adsk.core.GroupCommandInput.cast(
                inputs.itemById('bodyListGroup'))
            children = group.children if group else None

            raw_names = []
            for i, body in enumerate(bodies):
                if children and i < children.count:
                    inp = children.itemById('saveName_{}'.format(i))
                    raw = inp.value if inp else body.name
                else:
                    raw = body.name
                raw_names.append(_safe_filename(raw))

            name_count = {}
            for n in raw_names:
                name_count[n] = name_count.get(n, 0) + 1

            name_idx    = {}
            export_jobs = []        # list of (body, filepath)
            for body, n in zip(bodies, raw_names):
                idx = name_idx.get(n, 0) + 1
                name_idx[n] = idx
                fname = '{} ({})'.format(n, idx) if name_count[n] > 1 else n
                export_jobs.append(
                    (body, os.path.join(output_dir, fname + '.stl')))

            # ── Phase 1: Remove existing files (always overwrite) ────────
            overwritten = 0
            for _body, fpath in export_jobs:
                if os.path.exists(fpath):
                    try:
                        os.remove(fpath)
                        overwritten += 1
                    except OSError:
                        pass

            # ── Phase 2: Export with progress bar ────────────────────────
            total     = len(export_jobs)
            cancelled = False

            prog = _ui.createProgressDialog()
            prog.cancelButtonText        = 'Cancel'
            prog.isBackgroundTranslucent = False
            prog.isCancelButtonShown     = True
            prog.show(CMD_NAME, 'Exporting %v of %m …',
                      0, total, 0)
            adsk.doEvents()

            exported = 0
            errors   = []

            for i, (body, fpath) in enumerate(export_jobs):
                adsk.doEvents()
                if prog.wasCancelled:
                    cancelled = True
                    break
                prog.progressValue = i
                adsk.doEvents()

                try:
                    opts = export_mgr.createSTLExportOptions(body, fpath)
                    opts.meshRefinement = quality
                    export_mgr.execute(opts)
                    exported += 1
                except Exception as ex:
                    errors.append('{}: {}'.format(body.name, ex))

            prog.progressValue = total
            adsk.doEvents()
            prog.hide()

            # ── Summary ──────────────────────────────────────────────────
            msg = ['Exported: {}'.format(exported)]
            if overwritten:
                msg.append('Overwritten: {}'.format(overwritten))
            if errors:
                msg.append('Errors:   {}'.format(len(errors)))
                msg.extend('  \u2022 {}'.format(e) for e in errors)
            if cancelled:
                msg.append('\nExport cancelled by user.')
            msg.append('\nLocation: {}'.format(output_dir))
            _ui.messageBox('\n'.join(msg), CMD_NAME)

        except Exception:
            _ui.messageBox('Export failed:\n{}'.format(traceback.format_exc()))


# ──────────────────────────────────────────────────────────────────────────────
# Add-in Entry Points
# ──────────────────────────────────────────────────────────────────────────────
def run(context):
    """Called by Fusion 360 when the add-in is started."""
    global _app, _ui
    try:
        _app = adsk.core.Application.get()
        _ui  = _app.userInterface

        # Remove any leftover definition from a previous session
        old = _ui.commandDefinitions.itemById(CMD_ID)
        if old:
            old.deleteMe()

        cmd_def = _ui.commandDefinitions.addButtonDefinition(
            CMD_ID, CMD_NAME, CMD_DESC)

        # Place the command in the ADD-INS toolbar panel
        panel = _ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        ctrl  = panel.controls.itemById(CMD_ID)
        if ctrl:
            ctrl.deleteMe()
        panel.controls.addCommand(cmd_def)

        handler = _OnCommandCreated()
        cmd_def.commandCreated.add(handler)
        _handlers.append(handler)

    except Exception:
        if _ui:
            _ui.messageBox(
                'Multi Mesh Export failed to start:\n{}'.format(
                    traceback.format_exc()))


def stop(context):
    """Called by Fusion 360 when the add-in is stopped."""
    try:
        panel = _ui.allToolbarPanels.itemById('SolidScriptsAddinsPanel')
        if panel:
            ctrl = panel.controls.itemById(CMD_ID)
            if ctrl:
                ctrl.deleteMe()

        defn = _ui.commandDefinitions.itemById(CMD_ID)
        if defn:
            defn.deleteMe()

        _handlers.clear()

    except Exception:
        if _ui:
            _ui.messageBox(
                'Multi Mesh Export failed to stop:\n{}'.format(
                    traceback.format_exc()))
