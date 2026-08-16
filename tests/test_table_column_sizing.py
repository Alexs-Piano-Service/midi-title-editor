from types import SimpleNamespace

from aps_midi_prep_tool_app.main_window import MidiTitleWindow


class _Viewport:
    @staticmethod
    def width():
        return 1000


class _Header:
    @staticmethod
    def minimumSectionSize():
        return 20


class _Table:
    def __init__(self):
        self.widths = {0: 40, 2: 50, 3: 200, 4: 300, 6: 100}

    @staticmethod
    def rowCount():
        return 12

    @staticmethod
    def viewport():
        return _Viewport()

    @staticmethod
    def horizontalHeader():
        return _Header()

    @staticmethod
    def isColumnHidden(column):
        return column in {1, 5}

    def columnWidth(self, column):
        return self.widths.get(column, 0)

    def setColumnWidth(self, column, width):
        self.widths[column] = width


def test_batch_auto_fit_expands_title_column_to_full_file_list_width():
    table = _Table()
    window = SimpleNamespace(
        table=table,
        _is_adjusting_columns=False,
        _columns_content_fitted=False,
        _manual_column_widths={3: 200, 4: 300, 6: 100},
        _default_filename_column_width=lambda: 100,
        _minimum_title_column_width=lambda: 100,
        _minimum_user_resizable_column_width=lambda _column: 80,
    )

    MidiTitleWindow._auto_fit_table_columns_after_batch_change(window)

    visible_width = sum(table.widths[column] for column in (0, 2, 3, 4, 6))
    assert visible_width == table.viewport().width()
    assert table.widths[4] == 610
    assert window._columns_content_fitted is True
