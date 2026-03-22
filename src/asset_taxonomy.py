"""Asset classification taxonomy for JPX derivative underlyings.

Maps Japanese underlying_name values to structured categories
for systematic commodity analysis and dashboard organization.
"""

from __future__ import annotations


ASSET_TAXONOMY: dict[str, dict[str, str]] = {
    # ─── Energy: Power ───
    "電力(東・ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Base (Monthly)",
        "display_ja": "電力(東・ベース)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(東・日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Peak (Monthly)",
        "display_ja": "電力(東・日中)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(東・週間ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Base (Weekly)",
        "display_ja": "電力(東・週間ベース)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(東・週間日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Peak (Weekly)",
        "display_ja": "電力(東・週間日中)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(東・年度ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Base (Yearly)",
        "display_ja": "電力(東・年度ベース)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(東・年度日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power East Peak (Yearly)",
        "display_ja": "電力(東・年度日中)",
        "unit": "JPY/kWh",
        "region": "east",
    },
    "電力(西・ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Base (Monthly)",
        "display_ja": "電力(西・ベース)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    "電力(西・日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Peak (Monthly)",
        "display_ja": "電力(西・日中)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    "電力(西・週間ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Base (Weekly)",
        "display_ja": "電力(西・週間ベース)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    "電力(西・週間日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Peak (Weekly)",
        "display_ja": "電力(西・週間日中)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    "電力(西・年度ベース)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Base (Yearly)",
        "display_ja": "電力(西・年度ベース)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    "電力(西・年度日中)": {
        "category": "energy",
        "subcategory": "power",
        "display_en": "Power West Peak (Yearly)",
        "display_ja": "電力(西・年度日中)",
        "unit": "JPY/kWh",
        "region": "west",
    },
    # ─── Energy: Oil & Gas ───
    "ドバイ原油": {
        "category": "energy",
        "subcategory": "crude",
        "display_en": "Dubai Crude Oil",
        "display_ja": "ドバイ原油",
        "unit": "JPY/kl",
    },
    "LNG(プラッツJKM)": {
        "category": "energy",
        "subcategory": "lng",
        "display_en": "LNG (Platts JKM)",
        "display_ja": "LNG(プラッツJKM)",
        "unit": "USD/MMBtu",
    },
    "バージガソリン": {
        "category": "energy",
        "subcategory": "petroleum",
        "display_en": "Barge Gasoline",
        "display_ja": "バージガソリン",
        "unit": "JPY/kl",
    },
    "バージ灯油": {
        "category": "energy",
        "subcategory": "petroleum",
        "display_en": "Barge Kerosene",
        "display_ja": "バージ灯油",
        "unit": "JPY/kl",
    },
    "バージ軽油": {
        "category": "energy",
        "subcategory": "petroleum",
        "display_en": "Barge Gas Oil",
        "display_ja": "バージ軽油",
        "unit": "JPY/kl",
    },
    "中京ガソリン": {
        "category": "energy",
        "subcategory": "petroleum",
        "display_en": "Chukyo Gasoline",
        "display_ja": "中京ガソリン",
        "unit": "JPY/kl",
    },
    "中京灯油": {
        "category": "energy",
        "subcategory": "petroleum",
        "display_en": "Chukyo Kerosene",
        "display_ja": "中京灯油",
        "unit": "JPY/kl",
    },
    # ─── Precious Metals ───
    "金": {
        "category": "metals",
        "subcategory": "precious",
        "display_en": "Gold",
        "display_ja": "金",
        "unit": "JPY/g",
    },
    "金先物": {
        "category": "metals",
        "subcategory": "precious",
        "display_en": "Gold Futures",
        "display_ja": "金先物",
        "unit": "JPY/g",
    },
    "銀": {
        "category": "metals",
        "subcategory": "precious",
        "display_en": "Silver",
        "display_ja": "銀",
        "unit": "JPY/10g",
    },
    "白金": {
        "category": "metals",
        "subcategory": "precious",
        "display_en": "Platinum",
        "display_ja": "白金",
        "unit": "JPY/g",
    },
    "パラジウム": {
        "category": "metals",
        "subcategory": "precious",
        "display_en": "Palladium",
        "display_ja": "パラジウム",
        "unit": "JPY/g",
    },
    # ─── Industrial ───
    "ゴム(RSS3)": {
        "category": "industrial",
        "subcategory": "rubber",
        "display_en": "Rubber (RSS3)",
        "display_ja": "ゴム(RSS3)",
        "unit": "JPY/kg",
    },
    "ゴム(TSR20)": {
        "category": "industrial",
        "subcategory": "rubber",
        "display_en": "Rubber (TSR20)",
        "display_ja": "ゴム(TSR20)",
        "unit": "JPY/kg",
    },
    "上海天然ゴム": {
        "category": "industrial",
        "subcategory": "rubber",
        "display_en": "Shanghai Natural Rubber",
        "display_ja": "上海天然ゴム",
        "unit": "JPY/kg",
    },
    # ─── Agriculture ───
    "とうもろこし": {
        "category": "agriculture",
        "subcategory": "grains",
        "display_en": "Corn",
        "display_ja": "とうもろこし",
        "unit": "JPY/t",
    },
    "大豆": {
        "category": "agriculture",
        "subcategory": "grains",
        "display_en": "Soybeans",
        "display_ja": "大豆",
        "unit": "JPY/t",
    },
    "小豆": {
        "category": "agriculture",
        "subcategory": "grains",
        "display_en": "Azuki Beans",
        "display_ja": "小豆",
        "unit": "JPY/bag",
    },
    # ─── Equity Indices ───
    "日経225": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "Nikkei 225",
        "display_ja": "日経225",
        "unit": "JPY",
    },
    "日経平均VI": {
        "category": "equity",
        "subcategory": "volatility",
        "display_en": "Nikkei VI",
        "display_ja": "日経平均VI",
        "unit": "pts",
    },
    "日経平均気候変動指数": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "Nikkei Climate Change Index",
        "display_ja": "日経平均気候変動指数",
        "unit": "JPY",
    },
    "日経平均配当": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "Nikkei Average Dividend",
        "display_ja": "日経平均配当",
        "unit": "JPY",
    },
    "TOPIX": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "TOPIX",
        "display_ja": "TOPIX",
        "unit": "pts",
    },
    "TOPIX Core30": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "TOPIX Core30",
        "display_ja": "TOPIX Core30",
        "unit": "pts",
    },
    "JPX日経400": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "JPX-Nikkei 400",
        "display_ja": "JPX日経400",
        "unit": "pts",
    },
    "JPXプライム150指数": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "JPX Prime 150",
        "display_ja": "JPXプライム150指数",
        "unit": "pts",
    },
    "東証REIT": {
        "category": "equity",
        "subcategory": "reit",
        "display_en": "TSE REIT Index",
        "display_ja": "東証REIT",
        "unit": "pts",
    },
    "東証グロース市場250指数": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "TSE Growth 250",
        "display_ja": "東証グロース市場250指数",
        "unit": "pts",
    },
    "東証銀行": {
        "category": "equity",
        "subcategory": "sector",
        "display_en": "TSE Banks",
        "display_ja": "東証銀行",
        "unit": "pts",
    },
    "NYダウ": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "Dow Jones",
        "display_ja": "NYダウ",
        "unit": "JPY",
    },
    "FTSE中国50": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "FTSE China 50",
        "display_ja": "FTSE中国50",
        "unit": "pts",
    },
    "FTSE JPXネットゼロ・ジャパン500指数": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "FTSE JPX Net Zero Japan 500",
        "display_ja": "FTSE JPXネットゼロ・ジャパン500指数",
        "unit": "pts",
    },
    "S&P/JPX 500ESGスコア・ティルト指数": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "S&P/JPX 500 ESG Score Tilt",
        "display_ja": "S&P/JPX 500ESGスコア・ティルト指数",
        "unit": "pts",
    },
    "台湾加権": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "Taiwan TAIEX",
        "display_ja": "台湾加権",
        "unit": "pts",
    },
    "RNP": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "RNP",
        "display_ja": "RNP",
        "unit": "pts",
    },
    "CMEPI": {
        "category": "equity",
        "subcategory": "index",
        "display_en": "CME Petro Index",
        "display_ja": "CMEPI",
        "unit": "pts",
    },
    # ─── Bonds ───
    "長期国債": {
        "category": "bond",
        "subcategory": "jgb",
        "display_en": "Long-term JGB",
        "display_ja": "長期国債",
        "unit": "JPY",
    },
    "長期国債先物": {
        "category": "bond",
        "subcategory": "jgb",
        "display_en": "Long-term JGB Futures",
        "display_ja": "長期国債先物",
        "unit": "JPY",
    },
    "中期国債": {
        "category": "bond",
        "subcategory": "jgb",
        "display_en": "Medium-term JGB",
        "display_ja": "中期国債",
        "unit": "JPY",
    },
    "超長期国債": {
        "category": "bond",
        "subcategory": "jgb",
        "display_en": "Super Long-term JGB",
        "display_ja": "超長期国債",
        "unit": "JPY",
    },
    # ─── Interest Rate ───
    "無担保コールO/N物レート": {
        "category": "rate",
        "subcategory": "call",
        "display_en": "Unsecured Call O/N Rate",
        "display_ja": "無担保コールO/N物レート",
        "unit": "%",
    },
    # ─── FX ───
    "米ドル": {
        "category": "fx",
        "subcategory": "currency",
        "display_en": "USD/JPY",
        "display_ja": "米ドル",
        "unit": "JPY",
    },
}

# ─── Category metadata for display ───
CATEGORY_META = {
    "energy": {"display_en": "Energy", "display_ja": "エネルギー", "icon": "⚡", "color": "#F59E0B"},
    "metals": {"display_en": "Metals", "display_ja": "貴金属", "icon": "🥇", "color": "#FFD700"},
    "industrial": {"display_en": "Industrial", "display_ja": "産業素材", "icon": "🏭", "color": "#64748B"},
    "agriculture": {"display_en": "Agriculture", "display_ja": "農産物", "icon": "🌾", "color": "#22C55E"},
    "equity": {"display_en": "Equity", "display_ja": "株式指数", "icon": "📈", "color": "#3B82F6"},
    "bond": {"display_en": "Bonds", "display_ja": "国債", "icon": "🏛️", "color": "#8B5CF6"},
    "rate": {"display_en": "Interest Rates", "display_ja": "金利", "icon": "🏦", "color": "#06B6D4"},
    "fx": {"display_en": "FX", "display_ja": "為替", "icon": "💱", "color": "#EC4899"},
}

# Commodity categories (non-equity, non-bond) for the chart pack
COMMODITY_CATEGORIES = ["energy", "metals", "industrial", "agriculture"]


def get_assets_by_category(category: str) -> dict[str, dict[str, str]]:
    """Get all assets belonging to a given category."""
    return {
        name: info
        for name, info in ASSET_TAXONOMY.items()
        if info["category"] == category
    }


def get_assets_by_subcategory(subcategory: str) -> dict[str, dict[str, str]]:
    """Get all assets belonging to a given subcategory."""
    return {
        name: info
        for name, info in ASSET_TAXONOMY.items()
        if info["subcategory"] == subcategory
    }


def get_category_for_underlying(underlying_name: str) -> str | None:
    """Get the category for a given underlying_name. Returns None if not found."""
    info = ASSET_TAXONOMY.get(underlying_name)
    return info["category"] if info else None


def get_display_name(underlying_name: str, lang: str = "en") -> str:
    """Get display name for a given underlying. Falls back to original name."""
    info = ASSET_TAXONOMY.get(underlying_name)
    if info:
        return info[f"display_{lang}"]
    return underlying_name


def get_commodity_underlyings() -> list[str]:
    """Get all underlying names that are commodities (energy, metals, industrial, agriculture)."""
    return [
        name for name, info in ASSET_TAXONOMY.items()
        if info["category"] in COMMODITY_CATEGORIES
    ]


def get_non_power_commodity_underlyings() -> list[str]:
    """Get commodity underlyings excluding power futures."""
    return [
        name for name, info in ASSET_TAXONOMY.items()
        if info["category"] in COMMODITY_CATEGORIES and info["subcategory"] != "power"
    ]
