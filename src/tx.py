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
    Category.transport: {
        "WIND MOBILITY",
        "LIM*RIDE",
    },
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
        "אלקטרה מוצרי צריכה",
        "אמישרגז",
        "מי אביבים"
    },
    Category.savings: {
        "חסכון",
    },
    Category.donation: {
        "פעמונים",
        "חיים ביד",
        "מוסדות חב\"ד",
        "מכון מאיר",
        "עטרת",
        "מה יפו פעמי",
        "התורה והארץ",
        "גרעין יפו",
        "בית דוד בית שמש",
        "המרכז העולמי לחסד",
        "אסתר המלכה",
        "יד שרה",
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
        "חינוך",
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
        "רישיונות רכב",
        "חניוני",
        "רכב דוד",
    },
    Category.food: {
        "מכולת",
        "יוחננוף",
        "קפה עלית",
        "יינות ביתן",
        "חצות וחצי",
        "מחסני להב",
        "שופרסל",
        "רמי לוי",
        "מינימרקט",
        "מוצרי מזון",
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
        if self.amount < 0:
            return Category.income
        for kw, category in category_by_description.items():
            if kw in self.business:
                return category
        return Category.other
