import pandas as pd
import pytest

from excel_search.search import SearchCriterion, normalize_columns, search_dataframe


@pytest.fixture
def people() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Name": ["Anna Müller", "Bert Schmidt", "anna meier"],
            "Ort": ["Berlin", "Hamburg", "München"],
            "Code": ["A[1]", "B.2", "A[1]"],
        }
    )


def test_literal_search_does_not_treat_regex_characters_as_patterns(people: pd.DataFrame) -> None:
    result = search_dataframe(people, [SearchCriterion("[", column="Code")])

    assert result.index.tolist() == [0, 2]


def test_exact_search_is_case_insensitive_by_default(people: pd.DataFrame) -> None:
    result = search_dataframe(people, [SearchCriterion("ANNA MEIER", exact=True)])

    assert result.index.tolist() == [2]


def test_case_sensitive_search_is_consistent(people: pd.DataFrame) -> None:
    result = search_dataframe(
        people,
        [SearchCriterion("anna", column="Name")],
        case_sensitive=True,
    )

    assert result.index.tolist() == [2]


def test_all_and_any_combinations(people: pd.DataFrame) -> None:
    criteria = [SearchCriterion("Anna"), SearchCriterion("Berlin", column="Ort")]

    all_result = search_dataframe(people, criteria, match_mode="all")
    any_result = search_dataframe(people, criteria, match_mode="any")

    assert all_result.index.tolist() == [0]
    assert any_result.index.tolist() == [0, 2]


def test_unknown_column_is_reported(people: pd.DataFrame) -> None:
    with pytest.raises(ValueError, match="Unbekannte Spalte: Kudnenname"):
        search_dataframe(people, [SearchCriterion("Anna", column="Kudnenname")])


def test_empty_criteria_are_rejected(people: pd.DataFrame) -> None:
    with pytest.raises(ValueError, match="Mindestens ein Suchkriterium"):
        search_dataframe(people, [])


def test_headers_are_non_empty_unique_strings() -> None:
    dataframe = pd.DataFrame([[1, 2, 3]], columns=["Name", "Name", ""])

    result = normalize_columns(dataframe)

    assert result.columns.tolist() == ["Name", "Name (2)", "Spalte 3"]


def test_dates_are_searchable_in_display_format() -> None:
    dataframe = pd.DataFrame({"Datum": [pd.Timestamp("2026-08-21")]})

    result = search_dataframe(dataframe, [SearchCriterion("21.08.2026", exact=True)])

    assert result.index.tolist() == [0]
