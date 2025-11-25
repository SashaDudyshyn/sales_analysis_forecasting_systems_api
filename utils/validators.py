# utils/validators.py
from fastapi import HTTPException

def validate_column(col: str, name: str = "Стовпець") -> str:
    """Перевірка, чи колонка — це літера (A-Z, AA-ZZ тощо)"""
    col = col.strip().upper()
    if not col or not all(c.isalpha() for c in col):
        raise HTTPException(status_code=400, detail=f"{name}: некоректний формат стовпця '{col}'")
    return col


def validate_range(range_str: str) -> tuple[str, str]:
    """Перевірка діапазону E-G → ('E', 'G')"""
    range_str = range_str.strip().upper()
    if '-' not in range_str:
        raise HTTPException(status_code=400, detail=f"Діапазон має бути у форматі 'E-G', отримано: {range_str}")
    try:
        start, end = range_str.split('-', 1)
        validate_column(start, "Початок діапазону")
        validate_column(end, "Кінець діапазону")
        return start, end
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Невірний діапазон: {range_str}")


def validate_row(row: int, name: str) -> int:
    """Перевірка номера рядка"""
    if not isinstance(row, int) or row < 1:
        raise HTTPException(status_code=400, detail=f"{name} має бути цілим числом >= 1")
    return row


def validate_k(k: int) -> int:
    """Період згладжування"""
    if k < 1:
        raise HTTPException(status_code=400, detail="k має бути >= 1")
    return k


def validate_row_order(row_first: int, row_last: int):
    """rowFirstData <= rowLastData"""
    if row_first > row_last:
        raise HTTPException(status_code=400, detail="rowFirstData не може бути більше rowLastData")