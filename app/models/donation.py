from sqlalchemy import Column, ForeignKey, Integer, Text

from app.models.base import CharityDonationAbc


class Donation(CharityDonationAbc):
    user_id = Column(Integer, ForeignKey('user.id'))
    comment = Column(Text, nullable=True)

    def __repr__(self):
        return f'Donation({self.user_id}, {self.comment})'
