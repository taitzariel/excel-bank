import datetime
from dataclasses import dataclass
from enum import Enum
from typing import Any, Dict, Set


class Category(Enum):
    mortgage = "משכנתא"
    food = "אוכל"
    shvut = "שבות"
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
        "רב קו",
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
        "סלקום",
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
        "פעמוני",
        "חיים ביד",
        "מוסדות חב\"ד",
        "מכון מאיר",
        "עטרת",
        "מה יפו פעמי",
        "מאירים ביפו",
        "התורה והארץ",
        "גרעין יפו",
        "בית דוד בית שמש",
        "המרכז העולמי לחסד",
        "אסתר המלכה",
        "יד שרה",
        "ויצמן שולה",
        "עין-דרור",
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
        "אורות התורה",
        'ישיבת בנ"ע',
        "חינוך",
    },
    Category.atm: {
        "כספומט"
    },
    Category.mentoring: {
        "שר שלום",
        "שפר",
    },
    Category.fuel: {
        "פנגו",
        "דור אלון",
        "דלק",
        "סונול",
        "פז",
        "כלל חובה",
        "כלל אלמנטרי",
        "רישיונות רכב",
        "חניוני",
        "משרד התחבורה",
        "רכב דוד",
    },
    Category.shvut: {
        "מאיה אלגריסי",
        "צאלה קרני",
        "קרני צאלה",
    },
    Category.food: {
        "מכולת",
        "יוחננוף",
        "סופר דוש",
        "קפה עלית",
        "מגה קמעונאות",
        "יינות ביתן",
        "חצות וחצי",
        "מגה בעיר",
        "מחסני להב",
        "שופרסל",
        "אושר עד",
        "רמי לוי",
        "מינימרקט",
        "מוצרי מזון",
        "סיבוס",
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
