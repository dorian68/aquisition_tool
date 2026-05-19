from __future__ import annotations

from pathlib import Path

from .legacy_core import DASHBOARD_VBA_SOURCE as _LEGACY_VBA_SOURCE


_SETTINGS_NAVIGATION = """
Public Sub GoToSettings()
    SafeActivate "Settings"
End Sub

Public Sub ResetFilters()
    On Error GoTo Fail
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            If ws.FilterMode Then ws.ShowAllData
        End If
    Next ws
    ThisWorkbook.Worksheets("Dashboard").Activate
    MsgBox "Filters reset successfully.", vbInformation, "OptiQuant IA"
    Exit Sub
Fail:
    MsgBox "Reset filters failed: " & Err.Description, vbExclamation, "OptiQuant IA"
End Sub
"""

DASHBOARD_VBA_SOURCE = _LEGACY_VBA_SOURCE.replace(
    "Public Sub RefreshDashboard()",
    _SETTINGS_NAVIGATION.strip() + "\n\nPublic Sub RefreshDashboard()",
)


def write_vba_macro_source(path: str | Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(DASHBOARD_VBA_SOURCE.strip() + "\n", encoding="utf-8")
    return path
