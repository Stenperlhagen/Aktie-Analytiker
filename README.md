# EquityLab 3 – automatisk svensk bolagsanalys

## Funktioner
- Sök bland ett brett, manuellt kurerat urval svenska storbolag, inklusive OMXS30-namn, andra stora bolag och Volati/Salix.
- Hämta automatiskt aktiekurs, omsättning, marginal, nettoskuld, kassaflöde, EPS och P/E från Yahoo Finance (där data finns).
- Visa fem års kurshistorik och rapporterade finansiella tabeller.
- Sök efter nyheter från de senaste sex månaderna via Google News RSS.
- Ladda upp och sök i PDF, DOCX, XLSX, CSV och TXT. Skannade PDF-filer kräver OCR och stöds inte.
- Beräkna och exportera värderingsscenarier.

## Enklast: publicera som riktig hemsida
1. Skapa ett privat eller publikt repository på GitHub och ladda upp `app.py` och `requirements.txt`.
2. Gå till https://share.streamlit.io/ och välj **Create app** / **Deploy**.
3. Välj GitHub-repository och filen `app.py`.
4. Efter publicering får du en https-adress som fungerar utan installation för besökaren.

OBS: Om appen är publikt tillgänglig bör du inte ladda upp konfidentiella dokument. Använd privat drift och åtkomstkontroll för känsliga filer.

## Lokalt alternativ
`pip install -r requirements.txt`
`streamlit run app.py`

## Begränsningar
Detta är en fungerande Streamlit-app, men den är inte publicerad av oss och kan därför inte öppnas via en färdig publik URL. Yahoo Finance är en inofficiell källa och kan sakna nyckeltal eller ge inkonsekventa valuta-/perioduppgifter. Listan över bolag är ett stort urval, inte en verifierad fullständig Large Cap-lista. Nyheter hämtas på begäran och är inte en garanterat fullständig lista över regulatoriska pressmeddelanden. Dokumenten är sökbara men appen gör inte automatisk AI-tolkning av dokumenten.
