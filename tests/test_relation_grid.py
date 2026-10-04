"""关系表粘贴解析：制表符、换行、CRLF、行尾换行、长短不一的行。"""
import pytest

from web_app.relation_grid import coerce_relation_grid, parse_pasted_table, relation_payload_from_grid


def test_parse_pasted_table_tabs_newlines_crlf_trailing_and_ragged():
    assert parse_pasted_table("a\tb\nc\td\n") == [["a", "b"], ["c", "d"]]
    assert parse_pasted_table("a\tb\r\nc\td\r\n") == [["a", "b"], ["c", "d"]]
    assert parse_pasted_table("a\tb\nc") == [["a", "b"], ["c"]]
    assert parse_pasted_table("a\tb\tc\nd") == [["a", "b", "c"], ["d"]]
    assert parse_pasted_table("a\tb\n") == [["a", "b"]]
    assert parse_pasted_table("a\tb\n\n") == [["a", "b"], [""]]
    assert parse_pasted_table("a\tb\n\nc") == [["a", "b"], [""], ["c"]]
    assert parse_pasted_table("") == []
    assert parse_pasted_table("\n") == [[""]]
    assert parse_pasted_table("only\t\tcell\r\n") == [["only", "", "cell"]]


def test_ragged_grid_becomes_relation_payload():
    pasted = (
        "考核环节\t占比\t考核方式\t课程目标1\t课程目标2\t小计\r\n"
        "平时考核\t0.3\t平时作业\t50%\t50%\t100%\r\n"
        "\t\t课堂表现\t20%\t80%\r\n"
        "期末考核\t0.7\t期末考试\t40%\n"
    )
    rows = parse_pasted_table(pasted)
    assert rows[2] == ["", "", "课堂表现", "20%", "80%"]
    assert rows[3] == ["期末考核", "0.7", "期末考试", "40%"]
    payload = relation_payload_from_grid(rows)
    assert payload["objectives_count"] == 2
    assert [link["name"] for link in payload["links"]] == ["平时考核", "期末考核"]
    assert payload["links"][0]["methods"][1]["name"] == "课堂表现"
    assert payload["links"][1]["methods"][0]["supports"]["课程目标2"] == 0
    # 平时 0.5+0.2 乘 0.3，期末 0.4 乘 0.7
    assert payload["objectives_total_weights"]["课程目标1"] == pytest.approx(0.49, abs=1e-6)


def test_coerce_relation_grid_rejects_bad_shape():
    assert coerce_relation_grid("a\tb\n") == [["a", "b"]]
    with pytest.raises(ValueError):
        coerce_relation_grid({"a": 1})
    with pytest.raises(ValueError):
        coerce_relation_grid([["ok"], "bad"])
