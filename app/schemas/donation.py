from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field, PositiveInt


class DonationBase(BaseModel):
    full_amount: PositiveInt
    comment: Optional[str]
    invested_amount: Optional[int] = Field(0, hidden=True)
    fully_invested: Optional[bool] = Field(False, hidden=True)
    close_date: Optional[datetime] = Field(None, hidden=True)

    class Config:
        orm_mode = True

        @staticmethod
        def schema_extra(schema: dict, _):
            props = {}
            for k, v in schema.get('properties', {}).items():
                if not v.get("hidden", False):
                    props[k] = v
            schema["properties"] = props


class DonationCreate(DonationBase):
    ...


class DonationUserDB(BaseModel):
    full_amount: int
    comment: Optional[str]
    id: int
    create_date: datetime

    class Config:
        orm_mode = True


class DonationSuperUserDB(BaseModel):
    full_amount: int
    comment: str
    id: int
    create_date: datetime
    user_id: int
    invested_amount: int
    fully_invested: bool

    class Config:
        orm_mode = True
