# models/excel_params.py
from pydantic import BaseModel, Field, field_validator, model_validator

class ExcelProcessParams(BaseModel):
    #Основні дані
    column_year: str = Field(default="B", pattern=r"^[A-Z]+$")
    column_month: str = Field(default="D", pattern=r"^[A-Z]+$")
    range_data: str = Field(default="G-J", pattern=r"^[A-Z]+-[A-Z]+$")
    row_title: int = Field(default=3, ge=1, le=100)
    row_first_data: int = Field(default=4, ge=2, le=1000)
    row_last_data: int = Field(default=38, ge=5, le=5000)
    k: int = Field(default=2, ge=0, le=10)

    #Аркуші
    sheet_stat: str = Field(default="Статистичні дані", min_length=1)
    sheet_factor: str = Field(default="Фактори впливу", min_length=1)

    #Фактори впливу
    factor_column_year: str = Field(default="B", pattern=r"^[A-Z]+$")
    factor_column_month: str = Field(default="C", pattern=r"^[A-Z]+$")
    factor_row_range_data: str = Field(default="E-F", pattern=r"^[A-Z]+-[A-Z]+$")
    factor_row_description: int = Field(default=3, ge=1, le=100)
    factor_row_type: int = Field(default=4, ge=1, le=100)
    factor_row_title: int = Field(default=5, ge=1, le=100)
    factor_row_first_data: int = Field(default=6, ge=2, le=1000)
    factor_row_last_data: int = Field(default=17, ge=6, le=5000)

    # Крос-перевірки через field_validator
    @field_validator("range_data", "factor_row_range_data")
    @classmethod
    def range_start_before_end(cls, v: str) -> str:
        start, end = v.upper().split("-")
        if start >= end:
            raise ValueError(f"Початкова колонка ({start}) має бути лівіше за кінцеву ({end})")
        return v.upper()

    @field_validator("row_first_data")
    @classmethod
    def first_data_after_title(cls, v: int, info) -> int:
        if "row_title" in info.data and v <= info.data["row_title"]:
            raise ValueError("row_first_data має бути нижче row_title")
        return v

    @field_validator("row_last_data")
    @classmethod
    def last_after_first(cls, v: int, info) -> int:
        if "row_first_data" in info.data and v < info.data["row_first_data"]:
            raise ValueError("row_last_data має бути ≥ row_first_data")
        return v

    @field_validator("factor_row_last_data")
    @classmethod
    def factor_last_after_first(cls, v: int, info) -> int:
        if "factor_row_first_data" in info.data and v < info.data["factor_row_first_data"]:
            raise ValueError("factor_row_last_data має бути ≥ factor_row_first_data")
        return v

    # Перевірка унікальності рядків метаданих факторів
    @model_validator(mode="after")
    def factor_metadata_rows_distinct(self):
        rows = [
            self.factor_row_description,
            self.factor_row_type,
            self.factor_row_title,
        ]
        if len(set(rows)) != len(rows):
            raise ValueError("Рядки Опис / Тип / Заголовок факторів не можуть збігатися")
        return self

    # нормалізація колонок до uppercase
    @model_validator(mode="before")
    @classmethod
    def uppercase_columns(cls, data: dict):
        column_fields = [
            "column_year",
            "column_month",
            "factor_column_year",
            "factor_column_month",
        ]
        for field in column_fields:
            if field in data and isinstance(data[field], str):
                data[field] = data[field].upper()
        if "range_data" in data:
            data["range_data"] = data["range_data"].upper()
        if "factor_row_range_data" in data:
            data["factor_row_range_data"] = data["factor_row_range_data"].upper()
        return data
