Aktywne filtry =
VAR __separator = " | "

/* ---------- FUNKCJA POMOCNICZA (pattern) ---------- */
/* Tworzy tekst: label + top3 + "+ X innych" */
VAR __MakeText =
    VAR __dummy = 0
    RETURN __dummy
/* (placeholder – DAX nie ma funkcji, więc pattern kopiujemy poniżej) */


/* ---------- Proces ---------- */
VAR _Proces =
VAR _selected = VALUES(dim_P360[Proces])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_P360[Proces]),
    BLANK(),
    "Proces: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_P360[Proces], ASC),
            dim_P360[Proces],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Grupa procesów ---------- */
VAR _Grupa =
VAR _selected = VALUES(dim_P360[Grupa procesów])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_P360[Grupa procesów]),
    BLANK(),
    "Grupa procesów: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_P360[Grupa procesów], ASC),
            dim_P360[Grupa procesów],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Makroregion ---------- */
VAR _Makroregion =
VAR _selected = VALUES(dim_baza_kadrowa[Makroregion])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_baza_kadrowa[Makroregion]),
    BLANK(),
    "Makroregion: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_baza_kadrowa[Makroregion], ASC),
            dim_baza_kadrowa[Makroregion],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Obszar ---------- */
VAR _Obszar =
VAR _selected = VALUES(dim_baza_kadrowa[Obszar])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_baza_kadrowa[Obszar]),
    BLANK(),
    "Obszar: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_baza_kadrowa[Obszar], ASC),
            dim_baza_kadrowa[Obszar],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Oddział ---------- */
VAR _Oddzial =
VAR _selected = VALUES(dim_baza_kadrowa[Nazwa oddziału])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_baza_kadrowa[Nazwa oddziału]),
    BLANK(),
    "Oddział: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_baza_kadrowa[Nazwa oddziału], ASC),
            dim_baza_kadrowa[Nazwa oddziału],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Przypisanie ---------- */
VAR _Przypisanie =
VAR _selected = VALUES(dim_baza_kadrowa[Przypisanie])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_baza_kadrowa[Przypisanie]),
    BLANK(),
    "Przypisanie: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_baza_kadrowa[Przypisanie], ASC),
            dim_baza_kadrowa[Przypisanie],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)

/* ---------- Pracownik ---------- */
VAR _Pracownik =
VAR _selected = VALUES(dim_baza_kadrowa[Numer osobowy])
VAR _count = COUNTROWS(_selected)
RETURN
IF(
    NOT ISFILTERED(dim_baza_kadrowa[Numer osobowy]),
    BLANK(),
    "Pracownik: "
        & CONCATENATEX(
            TOPN(3, _selected, dim_baza_kadrowa[Numer osobowy], ASC),
            dim_baza_kadrowa[Numer osobowy],
            ", "
        )
        & IF(_count > 3, " + " & (_count - 3) & " innych", "")
)


/* ---------- SKLEJANIE ---------- */
VAR _result =
    CONCATENATEX(
        {
            _Proces,
            _Grupa,
            _Makroregion,
            _Obszar,
            _Oddzial,
            _Przypisanie,
            _Pracownik
        },
        [Value],
        __separator
    )

RETURN
IF(
    _result = "",
    "Aktualnie nie stosujesz żadnych filtrów",
    "Aktywne filtry: " & _result
)
