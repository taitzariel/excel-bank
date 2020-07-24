import datetime
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Set


class Category(Enum):
    mortgage = "משכנתא"
    food = "אוכל"
    education = "חינוך"
    running_expenses = "שוטף"
    mentoring = "הדרכה"
    donation = "תרומה"
    tax = "מס"
    insurance = "ביטוח"
    atm = "כספומט"
    fuel = "דלק"
    savings = "חסכון"
    transport = "תחבורה"
    other = "אחר"
    income = "הכנסות"


descriptions_by_category: Dict[Category, Set[str]] = {
    Category.mortgage: {
        "משכנתא",
    },
    Category.tax: {
        "מסים",
    },
    Category.running_expenses: {
        "ועד",
        "אינטרנט",
        "חברת החשמל",
        "019",
        "פלאפון",
        "בזק",
        "אמישרגז",
    },
    Category.savings: {
        "חסכון",
    },
    Category.donation: {
        "פעמונים",
        "מוסדות חב\"ד",
        "מכון מאיר",
        "עטרת",
        "מה יפו פעמי",
        "התורה והארץ",
        "גרעין יפו",
        "בית דוד בית שמש",
        "המרכז העולמי לחסד",
    },
    Category.insurance: {
        "מכבי",
        "שירותי ברי",
        "ביטוח",
        "פניקס",
        "מגדל",
    },
    Category.education: {
        "אמונה",
    },
    Category.atm: {
        "כספומט"
    },
    Category.mentoring: {
        "שר שלום",
    },
    Category.fuel: {
        "פנגו",
        "פז",
        "כלל חובה",
        "כלל אלמנטרי",
    },
    Category.food: {
        "מכולת",
        "יינות ביתן",
        "שופרסל",
        "רמי לוי",
    },
}

category_by_description: Dict[str, Category] = {
    keyword: cat for cat, keywords in descriptions_by_category.items() for keyword in keywords
}


@dataclass
class Transaction:
    amount: Any
    business: str
    charge_date: datetime.datetime
    transaction_date: datetime.datetime
    details: str
    card: str
    notes: str
    transaction_sum: Any

    def __post_init__(self) -> None:
        self.category = self._compute_category()

    def _compute_category(self) -> Category:
        for kw, category in category_by_description.items():
            if kw in self.business:
                return category
        if self.amount < 0:
            return Category.income
        else:
            return Category.other
