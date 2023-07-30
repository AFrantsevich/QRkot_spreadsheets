from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Extra, Field, PositiveInt, root_validator


class CharityProjectBase(BaseModel):
    name: str = Field(..., min_length=1, max_length=100)
    description: str = Field(..., min_length=1)
    full_amount: PositiveInt
    invested_amount: Optional[int] = Field(0, hidden=True)
    fully_invested: Optional[bool] = Field(False, hidden=True)

    class Config:
        orm_mode = True
        extra = Extra.forbid

        @staticmethod
        def schema_extra(schema: dict, _):
            props = {}
            for k, v in schema.get('properties', {}).items():
                if not v.get("hidden", False):
                    props[k] = v
            schema["properties"] = props


class CharityProjectCreate(CharityProjectBase):
    close_date: Optional[datetime] = Field(None, hidden=True)


class CharityProjectUpdate(BaseModel):
    name: Optional[str]
    description: Optional[str]
    full_amount: Optional[PositiveInt]

    class Config:
        orm_mode = True
        extra = Extra.forbid

    @root_validator()
    def name_cannot_be_null(cls, values):
        for field in values.values():
            if field is not None and not field:
                raise ValueError(f'Поле {field} проекта не может быть пустым!')
        return values


class CharityProjectDB(BaseModel):
    name: str
    description: str
    id: int
    full_amount: Optional[int]
    invested_amount: Optional[int]
    fully_invested: Optional[bool]
    create_date: Optional[datetime]
    close_date: Optional[datetime]

    class Config:
        orm_mode = True
