// za pomocą miary możemy stworzyć animowane kropki, które zależnie od statusu będą świecić wybranym kolorem lub będą wybrane mrugać

Status_Animated_SVG = 
VAR Data_URL = "data:image/svg+xml;utf8,"
VAR SVG_Start =
    "<svg height='100' width='100' xmlns='http://www.w3.org/2000/svg'>"

VAR Kolor_Statusu = [Status kolor] // może być tutaj funkcja a nie jako oddzialna ze SWITCH i SELECTEDVALUE
VAR Status_ = [_Status] // lub funkcja z SELECTEDVALUE
VAR Czas_Animacji = "3s"

VAR Bez_Animacji =
    "<circle r='30' cx='50' cy='50' fill='" & Kolor_Statusu & "'/>"

VAR Animacja =
    "<style>
        @keyframes blink {
            0% {opacity:1;}
            50% {opacity:0;}
            100% {opacity:1;}
        }
        circle {animation: blink " & Czas_Animacji & " infinite;}
     </style>
     <circle r='30' cx='50' cy='50' fill='" & Kolor_Statusu & "'/>"

VAR Check =
    IF(Status_ = "Opóźnione", Animacja, Bez_Animacji)

VAR SVG_End = "</svg>"

RETURN
    Data_URL & SVG_Start & Check & SVG_End



// zmieniamy format miary na image URL/Adres URL obrazu

