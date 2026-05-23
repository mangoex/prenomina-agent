from app.tools_prenomina import (
    DEFAULT_UMA_DIARIA,
    _calculate_fondo_ahorro,
    _calculate_fondo_ahorro_cap,
    _calculate_sdi,
    _normalize_yes_no,
)


def test_normalize_yes_no_true_values():
    for value in ["Si", "Sí", "si", "true", "1", "yes"]:
        assert _normalize_yes_no(value) is True


def test_normalize_yes_no_false_values():
    for value in ["No", "", "false", "0", None]:
        assert _normalize_yes_no(value) is False


def test_fondo_ahorro_sin_fondo():
    result = _calculate_fondo_ahorro(
        cuenta_fondo_ahorro="No",
        sueldo_nominal=10000,
        fondo_ahorro_excel=0,
        factor=0.11,
        uma_diaria=DEFAULT_UMA_DIARIA,
        tope_mode="none",
    )

    assert result["fondo_ahorro_base_calc"] == 0
    assert result["fondo_ahorro_calc"] == 0


def test_fondo_ahorro_con_fondo_calcula_base():
    result = _calculate_fondo_ahorro(
        cuenta_fondo_ahorro="Si",
        sueldo_nominal=10000,
        fondo_ahorro_excel=1100,
        factor=0.11,
        uma_diaria=DEFAULT_UMA_DIARIA,
        tope_mode="none",
    )

    assert result["fondo_ahorro_base_calc"] == 1100
    assert result["fondo_ahorro_calc"] == 1100


def test_fondo_ahorro_tope_mensual_no_excede_tope():
    cap = _calculate_fondo_ahorro_cap(
        uma_diaria=DEFAULT_UMA_DIARIA,
        tope_mode="mensual",
    )["tope"]
    result = _calculate_fondo_ahorro(
        cuenta_fondo_ahorro="Si",
        sueldo_nominal=100000,
        fondo_ahorro_excel=0,
        factor=0.11,
        uma_diaria=DEFAULT_UMA_DIARIA,
        tope_mode="mensual",
    )

    assert result["fondo_ahorro_calc"] == cap
    assert result["fondo_ahorro_calc"] <= cap


def test_salario_diario_integrado():
    assert round(_calculate_sdi(500, 1.0493), 2) == 524.65
