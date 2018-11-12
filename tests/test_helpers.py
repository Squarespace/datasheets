from __future__ import print_function

import copy
import datetime as dt
import sys

import numpy as np
import pandas as pd
import pytest

from datasheets import helpers


@pytest.mark.parametrize("data, expected", [
    ([[1, 'foo', 3], [2, 3, 4], [3, 'bar', 5]], 2),
    ([[1, 'foo', 3], [2, 'bar', None], [3, 4, 5]], 2),
    ([[1, 'foo', 3], [None, None, None], [3, 4, 5]], 2),
    ([[1, 'foo', 3], [3, 4, 5], [None, None, None]], 1),
    ([[None, None, None], [3, 4, 5], [None, None, None]], 1),
    ([[None, None, None], [None, None, None], [None, None, None]], None),
    ([], None)
    ])
def test_find_max_nonempty_row(data, expected):
    assert expected == helpers._find_max_nonempty_row(data)


@pytest.mark.parametrize("item, expected", [
    ('foo', 'foo'),
    (2, 2),
    (1.23, 1.23),
    (None, None),
    (dt.date(2016, 1, 1), '2016-01-01'),
    (dt.time(10, 20, 30), '10:20:30'),
    (dt.datetime(2016, 1, 1, 10, 20, 30), '2016-01-01 10:20:30'),
    (np.nan, None)
])
def test_convert_nan_and_datelike_values(item, expected):
    assert [[expected]] == helpers._convert_nan_and_datelike_values([[item]])


@pytest.mark.parametrize("row, col, expected", [
    (1, 1, 'A1'), (100, 27, 'AA100'), (7, 200, 'GR7')])
def test_convert_cell_index_to_label(row, col, expected):
    assert expected == helpers.convert_cell_index_to_label(row, col)


@pytest.mark.parametrize("row, col", [(-1, 1), (1, -1)])
def test_convert_cell_index_to_label_out_of_bounds(row, col):
    with pytest.raises(ValueError) as err:
        helpers.convert_cell_index_to_label(row, col)
    assert err.match('Row and column values must be >= 1')


@pytest.mark.parametrize("label, expected", [
    ('A1', (1, 1)), ('AA100', (100, 27)), ('GR7', (7, 200))])
def test_convert_cell_label_to_index(label, expected):
    assert expected == helpers.convert_cell_label_to_index(label)


@pytest.mark.parametrize("label", ['1', 'AA'])
def test_convert_cell_label_to_index_not_parseable(label):
    with pytest.raises(ValueError) as err:
        helpers.convert_cell_label_to_index(label)
    assert err.match('Unable to parse user-provided label')


def test_convert_cell_label_to_index_not_str():
    with pytest.raises(ValueError) as err:
        helpers.convert_cell_label_to_index(1)
    assert err.match('Input must be a string')


@pytest.mark.parametrize("query, expected", [
    ('', ''),
    ('some text', 'some text'),
    ("Quote'd text", "Quote\\'d text"),
    ("Backslashe\\d text", "Backslashe\\\\d text"),
    ("QuotedBackslashe\\'d text", "QuotedBackslashe\\\\\\'d text"),
])
def test_escape_query(query, expected):
    assert helpers._escape_query(query) == expected


@pytest.mark.parametrize("array, expected", [
    ([], []),
    ([None, None, None], []),
    ([None, None, 1], [None, None, 1]),
    (['foo', None, dt.datetime(2016, 1, 1)], ['foo', None, dt.datetime(2016, 1, 1)])
])
def test_remove_trailing_nones(array, expected):
    assert expected == helpers._remove_trailing_nones(array)


@pytest.mark.parametrize("array, new_len, expected", [
    ([], 3, [None, None, None]),
    ([1], 3, [1, None, None]),
    (['foo', 5, 'bar'], 2, ['foo', 5]),
    (['foo', 5, 'bar'], 3, ['foo', 5, 'bar']),
])
def test_resize_row(array, new_len, expected):
    assert expected == helpers._resize_row(array, new_len)


class TestMakeListOfLists:
    data = [[letter*i for i in range(2, 5)] for letter in 'abc']
    columns = ['col' + str(i) for i in range(1, 4)]
    row_idx = pd.Index(list('abc'), name='myindex')
    col_multiidx_named = pd.MultiIndex.from_tuples([('a', 'foo'), ('a', 'bar'), ('b', 'foo')], names=['cidx0', 'cidx1'])
    col_multiidx_unnamed = pd.MultiIndex.from_tuples([('a', 'foo'), ('a', 'bar'), ('b', 'foo')])
    row_multiidx_named = pd.MultiIndex.from_tuples([('baz', 1), ('cod', 3), ('baz', 1)], names=['ridx0', 'ridx1'])
    row_multiidx_unnamed = pd.MultiIndex.from_tuples([('baz', 1), ('cod', 3), ('baz', 1)])

    df_noindex = pd.DataFrame(data=data, columns=columns)
    df_index = pd.DataFrame(data=data, columns=columns, index=row_idx)
    df_col_multiidx_named = pd.DataFrame(data=data, columns=col_multiidx_named)
    df_col_multiidx_unnamed = pd.DataFrame(data=data, columns=col_multiidx_unnamed)
    df_row_multiidx_named = pd.DataFrame(data=data, columns=columns, index=row_multiidx_named)
    df_row_multiidx_unnamed = pd.DataFrame(data=data, columns=columns, index=row_multiidx_unnamed)
    df_dual_multiidx_named = pd.DataFrame(data=data, columns=col_multiidx_named, index=row_multiidx_named)
    df_dual_multiidx_unnamed = pd.DataFrame(data=data, columns=col_multiidx_unnamed, index=row_multiidx_unnamed)

    def test_list_of_lists(self):
        data = [[None, 'foo'], [1, 'bar']]
        headers, values = helpers._make_list_of_lists(data, index=False)
        assert headers == [] and values == [[None, 'foo'], [1, 'bar']]

    def test_list_of_dicts(self):
        data = [
            {'name': 'bubbles', 'age': 3},
            {'name': 'fishsticks', 'age': 2.5},
            {'name': 'Mr. Whiskers', 'age': 6},
        ]

        headers, values = helpers._make_list_of_lists(data, index=False)
        assert len(headers) == 1
        assert set(headers[0]) == set(['name', 'age'])

        expected_values = [
            ['bubbles', 3],
            ['fishsticks', 2.5],
            ['Mr. Whiskers', 6],
        ]
        reversed_expected_values = [row[::-1] for row in expected_values]
        # We have to check both orderings since dict keys aren't ordered
        assert values == expected_values or values == reversed_expected_values

    def test_df_no_index_index_true(self):
        headers = [['index0'] + self.columns]
        values = [[i] + row for i, row in enumerate(self.data)]
        assert helpers._make_list_of_lists(self.df_noindex, index=True) == (headers, values)

    def test_df_no_index_index_false(self):
        expected = ([self.columns], self.data)
        assert helpers._make_list_of_lists(self.df_noindex, index=False) == expected

    def test_df_simple_row_index_index_true(self):
        headers = [['myindex'] + list(self.df_index.columns)]
        values = [[i] + j for i, j in zip(self.df_index.index, self.data)]
        assert helpers._make_list_of_lists(self.df_index, index=True) == (headers, values)

    def test_df_simple_row_index_index_false(self):
        headers = [list(self.df_index.columns)]
        values = self.data
        assert helpers._make_list_of_lists(self.df_index, index=False) == (headers, values)

    def test_df_unnamed_column_multiindex_index_true(self):
        headers = list(map(list, zip(('index0', 'index0'), *self.df_col_multiidx_unnamed.columns)))
        values = [[i] + row for i, row in enumerate(self.data)]
        assert helpers._make_list_of_lists(self.df_col_multiidx_unnamed, index=True) == (headers, values)

    def test_df_unnamed_column_multiindex_index_false(self):
        headers = list(map(list, zip(*self.df_col_multiidx_unnamed.columns)))
        values = self.data
        assert helpers._make_list_of_lists(self.df_col_multiidx_unnamed, index=False) == (headers, values)

    def test_df_named_column_multiindex_index_true(self):
        headers = list(map(list, zip(('index0', 'index0'), *self.df_col_multiidx_named.columns)))
        values = [[i] + row for i, row in enumerate(self.data)]
        assert helpers._make_list_of_lists(self.df_col_multiidx_named, index=True) == (headers, values)

    def test_df_named_column_multiindex_index_false(self):
        headers = list(map(list, zip(*self.df_col_multiidx_named.columns)))
        values = self.data
        assert helpers._make_list_of_lists(self.df_col_multiidx_named, index=False) == (headers, values)

    def test_df_unnamed_row_multiindex_index_true(self):
        headers = [['index0', 'index1'] + list(self.df_row_multiidx_unnamed.columns)]
        values = [list(i) + j for i, j in zip(self.df_row_multiidx_unnamed.index, self.data)]
        assert helpers._make_list_of_lists(self.df_row_multiidx_unnamed, index=True) == (headers, values)

    def test_df_unnamed_row_multiindex_index_false(self):
        headers = [list(self.df_row_multiidx_unnamed.columns)]
        values = self.data
        assert helpers._make_list_of_lists(self.df_row_multiidx_unnamed, index=False) == (headers, values)

    def test_df_named_row_multiindex_true(self):
        headers = [list(self.df_row_multiidx_named.index.names) + list(self.df_row_multiidx_named.columns)]
        values = [list(i) + j for i, j in zip(self.df_row_multiidx_named.index, self.data)]
        assert helpers._make_list_of_lists(self.df_row_multiidx_named, index=True) == (headers, values)

    def test_df_named_row_multiindex_false(self):
        headers = [list(self.df_row_multiidx_named.columns)]
        values = self.data
        assert helpers._make_list_of_lists(self.df_row_multiidx_named, index=False) == (headers, values)

    def test_df_unnamed_row_and_column_multiindexes_index_true(self):
        add_unnamed_entries = lambda row: ['index0', 'index1'] + list(row)
        headers = list(map(add_unnamed_entries, zip(*self.df_dual_multiidx_unnamed.columns)))
        values = [list(i) + j for i, j in zip(self.df_dual_multiidx_unnamed.index, self.data)]
        assert helpers._make_list_of_lists(self.df_dual_multiidx_unnamed, index=True) == (headers, values)

    def test_df_unnamed_row_and_column_multiindexes_index_false(self):
        headers = list(map(list, zip(*self.df_dual_multiidx_unnamed.columns)))
        values = self.data
        assert helpers._make_list_of_lists(self.df_dual_multiidx_unnamed, index=False) == (headers, values)

    def test_df_named_row_and_column_multiindexes_index_true(self):
        add_named_entries = lambda row: list(self.df_dual_multiidx_named.index.names) + list(row)
        headers = list(map(add_named_entries, zip(*self.df_dual_multiidx_named.columns)))
        values = [list(i) + j for i, j in zip(self.df_dual_multiidx_named.index, self.data)]
        assert helpers._make_list_of_lists(self.df_dual_multiidx_named, index=True) == (headers, values)

    def test_df_named_row_and_column_multiindexes_index_false(self):
        headers = list(map(list, zip(*self.df_dual_multiidx_named.columns)))
        values = self.data
        assert helpers._make_list_of_lists(self.df_dual_multiidx_named, index=False) == (headers, values)

    def test_value_error(self):
        wrong_data_type = dict(foo='bar')
        with pytest.raises(ValueError) as err:
            helpers._make_list_of_lists(wrong_data_type, index=False)
        assert err.match('Input data must be a pandas.DataFrame, a list of dicts, or a list of lists')
