settings = {
    "sheet_id": "123456789",
    "headers": {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36",
    },
    "day_mapping": {
        0: "週一",
        1: "週二",
        2: "週三",
        3: "週四",
        4: "週五",
        5: "週六",
        6: "週日",
    },
    "copy_only": {
        "散戶選擇權": {
            "start_row": 2,
            "copy_cell_data": {
                "start_column": 0,
                "end_column": 46,
            },
        }
    },
    "三大法人期貨": {
        "大台": {
            "url": "https://www.taifex.com.tw/cht/3/futContractsDateDown",
            "code": "TXF",
            "start_row": 4,
            "work_sheet": {
                "外資及陸資": "外資大台",
                "自營商": "自營商大台",
                "投信": "投信大台",
            },
            "copy_cell_data": {
                "start_column": 15,
                "end_column": 26,
            },
        },
        "小台": {
            "url": "https://www.taifex.com.tw/cht/3/futContractsDateDown",
            "code": "MXF",
            "start_row": 4,
            "work_sheet": {
                "外資及陸資": "外資小台",
                "自營商": "自營商小台",
            },
            "copy_cell_data": {
                "start_column": 15,
                "end_column": 26,
            },
        },
    },
    "選擇權": {
        "大台": {
            "url": "https://www.taifex.com.tw/cht/3/callsAndPutsDateDown",
            "code": "TXO",
            "start_row": 12,
            "work_sheet": {
                "外資及陸資_CALL": "外資OP買權",
                "外資及陸資_PUT": "外資OP賣權",
                "自營商_CALL": "自營商OP買權",
                "自營商_PUT": "自營商OP賣權",
            },
        }
    },
    "期貨每日交易行情": {
        "大台": {
            "url": "https://www.taifex.com.tw/cht/3/futDataDown",
            "code": "TX",
            "down_type": 1,
            "start_row": 4,
            "previous_hold": "大台期散戶!G6",
            "previous_transaction": "大台期散戶!E6",
            "work_sheet": "大台期散戶",
            "copy_cell_data": {
                "start_column": 8,
                "end_column": 43,
            },
        },
        "小台": {
            "url": "https://www.taifex.com.tw/cht/3/futDataDown",
            "code": "MTX",
            "down_type": 1,
            "start_row": 4,
            "previous_hold": "小台期散戶!G6",
            "previous_transaction": "小台期散戶!E6",
            "work_sheet": "小台期散戶",
            "copy_cell_data": {
                "start_column": 8,
                "end_column": 35,
            },
        },
    },
    "期貨大額交易人未沖銷部位": {
        "大台": {
            "url": "https://www.taifex.com.tw/cht/3/largeTraderFutDown",
            "code": "TX",
            "start_row": 1,
            "work_sheet": {
                "特法期貨": "臺股期貨",
            },
        },
    },
    "選擇權大額交易人未沖銷部位": {
        "大台": {
            "url": "https://www.taifex.com.tw/cht/3/largeTraderOptDown",
            "code": "TXO",
            "start_row": 3,
            "work_sheet": {
                "臺指買權": "特法OP買權",
                "臺指賣權": "特法OP賣權",
            },
        },
    },
    "玩股網": {
        "bank_mapping": {
            "bot": 2,  # 台銀
            "land": 1,  # 土銀
            "tcb": 0,  # 合庫
            "hncb": 7,  # 華南永昌
            "mega": 6,  # 兆豐銀
            "tbb": 3,  # 台企銀
            "chb": 4,  # 彰銀
            "first": 5,  # 第一金
        },
        "main_url": "https://www.wantgoo.com/stock/institutional-investors/three-trade-for-trading-amount",
        "日趨勢": {
            "url": "https://www.wantgoo.com/investrue/wtx&/daily-candlesticks?after={}",
            "start_row": 2,
            "copy_cell_data": {
                "start_column": 5,
                "end_column": 6,
            },
            "copy_cell_formula": {
                "start_column": 7,
                "end_column": 11,
            },
            "work_sheet": "日趨勢",
        },
        "買賣權比": {
            "url": "https://www.wantgoo.com/option/put-call-ratio-data",
            "start_row": 1,
            "copy_cell_data": {
                "start_column": 9,
                "end_column": 10,
            },
            "work_sheet": "P/C Ratio",
        },
        "八大行庫買賣動向": {
            "url": "https://www.wantgoo.com/stock/public-bank/trend-data?market=-1",
            "start_row": 2,
            "work_sheet": "八大行庫買賣動向",
            "copy_cell_data": {
                "start_column": 17,
                "end_column": 19,
            },
        },
        "url": {
            "big_8_buycall": "https://www.wantgoo.com/stock/public-bank/buy-sell-data?market=-1&orderBy=count&during=1&industry=",
        },
    },
}
