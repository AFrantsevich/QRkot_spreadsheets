from sqlalchemy import Column, String, Text

from app.models.base import CharityDonationAbc


class CharityProject(CharityDonationAbc):
    name = Column(String(100), unique=True, nullable=False)
    description = Column(Text, nullable=False)

    def __repr__(self):
        return f'CharityProject({self.name}, {self.description})'
