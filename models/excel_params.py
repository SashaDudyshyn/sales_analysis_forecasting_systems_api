# models/excel_params.py
from pydantic import BaseModel, field_validator, Field, model_validator
from typing import Optional

class ExcelProcessParams(BaseModel):
    column_year: str = Field(..., description="Стовпець з роком (наприклад, B)")
    column_month: str = Field(..., description="Стовпець з місяцем (наприклад, D)")
    range_data: str = Field(..., description="Діапазон даних (наприклад, E-G)")
    row_title: int = Field(..., ge=1)
    row_first_data: int = Field(..., ge=1)
    row_last_data: int = Field(..., ge=1)
    k: Optional[int] = Field(2, ge=1)

    # Валідація стовпців
    @field_validator("column_year", "column_month")
    @classmethod
    def validate_column(cls, v: str) -> str:
        v = v.strip().upper()
        if not v or not all(c.isalpha() for c in v):
            raise ValueError(f"Некоректний формат стовпця: {v}")
        return v

    # Валідація діапазону
    @field_validator("range_data")
    @classmethod
    def validate_range(cls, v: str) -> str:
        v = v.strip().upper()
        if "-" not in v:
            raise ValueError("Діапазон має бути у форматі 'E-G'")
        start, end = v.split("-", 1)
        if not (start and end and start.isalpha() and end.isalpha()):
            raise ValueError(f"Невірний діапазон: {v}")
        return v

    # Перевірка порядку рядків — через model_validator
    @model_validator(mode="after")
    def check_row_order(self):
        if self.row_first_data > self.row_last_data:
            raise ValueError("row_first_data не може бути більше row_last_data")
        return self