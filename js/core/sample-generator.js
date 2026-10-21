/**
 * Sample Generator - Creates realistic multi-language Code of Ethics dataset.
 * Supports German (2024 & 2026) + 10 target languages (EN, CS, FR, ES, IT, PL, NL, SK, HU, RO).
 * Universal module: works in Node.js and browser.
 */
(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define([], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.SampleGenerator = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function getSampleData() {
    return [
      {
        'Artikel-Nr.': 'Art. 01',
        'Bereich': 'Präambel & Grundlagen',
        'DE 2024 (Referenz)': 'Dieser Verhaltenskodex definiert die grundlegenden ethischen und rechtlichen Standards, die für alle Mitarbeiter und Führungskräfte des Unternehmens weltweit verbindlich sind.',
        'DE 2026 (Referenz)': 'Dieser Verhaltenskodex definiert die grundlegenden ethischen und rechtlichen Standards, die für alle Mitarbeiter und Führungskräfte des Unternehmens weltweit verbindlich sind.',
        'EN 2024': 'This Code of Conduct defines the fundamental ethical and legal standards that are binding for all employees and executives of the company worldwide.',
        'CS 2024': 'Tento etický kodex definuje základní etické a právní standardy, které jsou závazné pro všechny zaměstnance a vedoucí pracovníky společnosti po celém světě.',
        'FR 2024': 'Le présent code de conduite définit les normes éthiques et juridiques fondamentales qui lient l\'ensemble des collaborateurs et dirigeants de l\'entreprise dans le monde entier.',
        'ES 2024': 'Este Código de Conducta define las normas éticas y legales fundamentales que son vinculantes para todos los empleados y directivos de la empresa en todo el mundo.',
        'IT 2024': 'Il presente Codice di Condotta definisce gli standard etici e legali fondamentali vincolanti per tutti i dipendenti e i dirigenti dell\'azienda in tutto il mondo.',
        'PL 2024': 'Niniejszy Kodeks Postępowania określa podstawowe standardy etyczne i prawne, które są wiążące dla wszystkich pracowników i kadry zarządzającej firmy na całym świecie.',
        'NL 2024': 'Deze gedragscode definieert de fundamentele ethische en wettelijke normen die bindend zijn voor alle medewerkers en leidinggevenden van het bedrijf wereldwijd.',
        'SK 2024': 'Tento etický kódex definuje základné etické a právne štandardy, ktoré sú záväzné pre všetkých zamestnancov a vedúcich pracovníkov spoločnosti na celom svete.',
        'HU 2024': 'Ez a magatartási kódex meghatározza azokat az alapvető etikai és jogi normákat, amelyek a vállalat valamennyi munkavállalójára és vezetőjére világszerte kötelező érvényűek.',
        'RO 2024': 'Prezentul Cod de conduită definește standardele etice și juridice fundamentale obligatorii pentru toți angajații și conducerea companiei din întreaga lume.'
      },
      {
        'Artikel-Nr.': 'Art. 02',
        'Bereich': 'Gesetzestreue',
        'DE 2024 (Referenz)': 'Wir verpflichten uns zur uneingeschränkten Einhaltung aller anwendbaren Gesetze, Verordnungen und internen Richtlinien in allen Ländern, in denen wir tätig sind.',
        'DE 2026 (Referenz)': 'Wir verpflichten uns zur uneingeschränkten Einhaltung aller anwendbaren Gesetze, Verordnungen und internen Richtlinien in allen Ländern, in denen wir tätig sind.',
        'EN 2024': 'We commit to strict compliance with all applicable laws, regulations, and internal policies in all countries in which we operate.',
        'CS 2024': 'Zavazujeme se k bezvýhradnému dodržování všech platných zákonů, předpisů a interních směrnic ve všech zemích, v nichž působíme.',
        'FR 2024': 'Nous nous engageons à respecter scrupuleusement l\'ensemble des lois, réglementations et directives internes applicables dans tous les pays où nous exerçons nos activités.',
        'ES 2024': 'Nos comprometemos al cumplimiento irrestricto de todas las leyes, regulaciones y directrices internas aplicables en todos los países donde operamos.',
        'IT 2024': 'Ci impegniamo al rispetto incondizionato di tutte le leggi, i regolamenti e le direttive interne applicabili in tutti i paesi in cui operiamo.',
        'PL 2024': 'Zobowiązujemy się do bezwzględnego przestrzegania wszystkich obowiązujących przepisów prawa, rozporządzeń i wewnętrznych wytycznych we wszystkich krajach naszej działalności.',
        'NL 2024': 'Wij verbinden ons tot de strikte naleving van alle toepasselijke wetten, voorschriften en interne richtlijnen in alle landen waarin wij actief zijn.',
        'SK 2024': 'Zaväzujeme sa k bezvýhradnému dodržiavaniu všetkých platných zákonov, predpisov a interných smerníc vo všetkých krajinách, kde pôsobíme.',
        'HU 2024': 'Kötelezettséget vállalunk az összes alkalmazandó törvény, rendelet és belső irányelv maradéktalan betartására mindazon országokban, ahol tevékenységet folytatunk.',
        'RO 2024': 'Ne angajăm să respectăm necondiționat toate legile, reglementările și politicile interne aplicabile în toate țările în care ne desfășurăm activitatea.'
      },
      {
        'Artikel-Nr.': 'Art. 03',
        'Bereich': 'Chancengleichheit',
        'DE 2024 (Referenz)': 'Jegliche Diskriminierung aufgrund von Geschlecht, Rasse, ethnischer Herkunft, Religion oder Behinderung wird nicht toleriert.',
        'DE 2026 (Referenz)': 'Jegliche Diskriminierung aufgrund von Geschlecht, Geschlechtsidentität, ethnischer Herkunft, Religion, Alter oder Behinderung wird ausnahmslos nicht toleriert.',
        'EN 2024': 'Any discrimination based on gender, race, ethnic origin, religion, or disability will not be tolerated.',
        'CS 2024': 'Jakákoli diskriminace na základě pohlaví, rasy, etnického původu, náboženství nebo zdravotního postižení je netolerována.',
        'FR 2024': 'Toute discrimination fondée sur le sexe, la race, l\'origine ethnique, la religion ou le handicap ne sera aucunement tolérée.',
        'ES 2024': 'No se tolerará ninguna discriminación por motivos de género, raza, origen étnico, religión o discapacidad.',
        'IT 2024': 'Non sarà tollerata alcuna discriminazione basata su sesso, razza, origine etnica, religione o disabilità.',
        'PL 2024': 'Wszelka dyskryminacja ze względu na płeć, rasę, pochodzenie etniczne, religię lub niepełnosprawność nie będzie tolerowana.',
        'NL 2024': 'Elke vorm van discriminatie op grond van geslacht, ras, etnische afkomst, religie of handicap wordt niet getolereerd.',
        'SK 2024': 'Akákoľvek diskriminácia na základe pohlavia, rasy, etnického pôvodu, náboženstva alebo zdravotného postihnutia sa netoleruje.',
        'HU 2024': 'Semmilyen nemen, fajon, etnikai származáson, valláson vagy fogyatékosságon alapuló megkülönböztetést nem tűrünk el.',
        'RO 2024': 'Nu se tolerează nicio discriminare bazată pe sex, rasă, origine etnică, religie sau dizabilitate.'
      },
      {
        'Artikel-Nr.': 'Art. 04',
        'Bereich': 'Geschenke & Bewirtung',
        'DE 2024 (Referenz)': 'Geschenke und Einladungen dürfen nur angenommen werden, wenn ihr Wert 35 Euro nicht übersteigt und sie den üblichen geschäftlichen Gepflogenheiten entsprechen.',
        'DE 2026 (Referenz)': 'Geschenke und Einladungen dürfen nur angenommen werden, wenn ihr Wert 50 Euro nicht übersteigt und sie vorab über das Compliance-Portal freigegeben wurden.',
        'EN 2024': 'Gifts and invitations may only be accepted if their value does not exceed 35 euros and they comply with customary business practices.',
        'CS 2024': 'Dary a pozvání lze přijmout pouze tehdy, pokud jejich hodnota nepřesahuje 35 eur a odpovídají běžným obchodním zvyklostem.',
        'FR 2024': 'Les cadeaux et invitations ne peuvent être acceptés que si leur valeur ne dépasse pas 35 euros et qu\'ils sont conformes aux usages commerciaux habituels.',
        'ES 2024': 'Los obsequios e invitaciones solo podrán aceptarse si su valor no supera los 35 euros y se ajustan a las prácticas comerciales habituales.',
        'IT 2024': 'Regali e inviti possono essere accettati solo se il loro valore non supera i 35 euro e sono conformi alle consuete pratiche commerciali.',
        'PL 2024': 'Prezenty i zaproszenia można przyjmować tylko wtedy, gdy ich wartość nie przekracza 35 euro i są one zgodne ze zwyczajowymi praktykami biznesowymi.',
        'NL 2024': 'Geschenken en uitnodigingen mogen alleen worden geaccepteerd als de waarde ervan niet hoger is dan 35 euro en zij voldoen aan de gebruikelijke zakelijke gebruiken.',
        'SK 2024': 'Dary a pozvania je možné prijať iba vtedy, ak ich hodnota nepresahuje 35 eur a zodpovedajú bežným obchodným zvyklostiam.',
        'HU 2024': 'Ajándékok és meghívások csak akkor fogadhatók el, ha értékük nem haladja meg a 35 eurót, és megfelelnek a szokásos üzleti gyakorlatnak.',
        'RO 2024': 'Cadourile și invitațiile pot fi acceptate numai dacă valoarea acestora nu depășește 35 de euro și respectă practicile de afaceri uzuale.'
      },
      {
        'Artikel-Nr.': 'Art. 05',
        'Bereich': 'Interessenkonflikte',
        'DE 2024 (Referenz)': 'Geschäftliche Entscheidungen müssen ausschließlich im besten Interesse des Unternehmens getroffen werden, frei von persönlichen oder familiären Vorteilen.',
        'DE 2026 (Referenz)': 'Geschäftliche Entscheidungen müssen ausschließlich im besten Interesse des Unternehmens getroffen werden, frei von persönlichen oder familiären Vorteilen.',
        'EN 2024': 'Business decisions must be made solely in the best interest of the company, free from personal or family advantages.',
        'CS 2024': 'Obchodní rozhodnutí musí být činěna výhradně v nejlepším zájmu společnosti, bez osobních či rodinných výhod.',
        'FR 2024': 'Les décisions commerciales doivent être prises exclusivement dans le meilleur intérêt de l\'entreprise, exemptes de tout avantage personnel ou familial.',
        'ES 2024': 'Las decisiones comerciales deben tomarse exclusivamente en el mejor interés de la empresa, libres de ventajas personales o familiares.',
        'IT 2024': 'Le decisioni aziendali devono essere prese esclusivamente nel migliore interesse dell\'azienda, prive di vantaggi personali o familiari.',
        'PL 2024': 'Decyzje biznesowe muszą być podejmowane wyłącznie w najlepszym interesie firmy, wolne od korzyści osobistych lub rodzinnych.',
        'NL 2024': 'Zakelijke beslissingen moeten uitsluitend worden genomen in het beste belang van het bedrijf, vrij van persoonlijke of familiale voordelen.',
        'SK 2024': 'Obchodné rozhodnutia sa musia prijímať výlučne v najlepšom záujme spoločnosti, bez osobných či rodinných výhod.',
        'HU 2024': 'Az üzleti döntéseket kizárólag a vállalat legjobb érdekében kell meghozni, mentesen a személyes vagy családi előnyöktől.',
        'RO 2024': 'Deciziile de afaceri trebuie luate exclusiv în interesul superior al companiei, fără avantaje personale sau de familie.'
      },
      {
        'Artikel-Nr.': 'Art. 06',
        'Bereich': 'Antikorruption',
        'DE 2024 (Referenz)': 'Bestechung und Schmiergeldzahlungen in jeglicher Form an Amtsträger oder private Geschäftspartner sind strengstens verboten.',
        'DE 2026 (Referenz)': 'Bestechung und Schmiergeldzahlungen in jeglicher Form an Amtsträger oder private Geschäftspartner sind strengstens verboten.',
        'EN 2024': 'Bribery and kickback payments in any form to public officials or private business partners are strictly prohibited.',
        'CS 2024': 'Úplatkářství a jakékoliv formy provizí či úplatků úředním osobám nebo soukromým obchodním partnerům jsou přísně zakázány.',
        'FR 2024': 'La corruption et les pots-de-vin sous quelque forme que ce soit à des agents publics ou à des partenaires commerciaux privés sont strictement interdits.',
        'ES 2024': 'El soborno y los pagos de comisiones ilegales en cualquier forma a funcionarios públicos o socios comerciales privados están estrictamente prohibidos.',
        'IT 2024': 'La corruzione e i pagamenti di tangenti in qualsiasi forma a pubblici ufficiali o partner commerciali privati sono rigorosamente vietati.',
        'PL 2024': 'Łapownictwo i płatności korupcyjne w jakiejkolwiek formie na rzecz urzędników lub prywatnych partnerów biznesowych są surowo zabronione.',
        'NL 2024': 'Omkoping en smeergelden in welke vorm dan ook aan ambtenaren of particuliere zakenpartners zijn ten strengste verboden.',
        'SK 2024': 'Úplatkárstvo a akékoľvek formy provízií či úplatkov úradným osobám alebo súkromným obchodným partnerom sú prísne zakázané.',
        'HU 2024': 'A vesztegetés és a kenőpénzek bármilyen formája köztisztviselők vagy magán üzleti partnerek felé szigorúan tilos.',
        'RO 2024': 'Darea de mită și plățile de facilitare sub orice formă către funcționari publici sau parteneri de afaceri privați sunt strict interzise.'
      },
      {
        'Artikel-Nr.': 'Art. 07',
        'Bereich': 'Datenschutz & Vertraulichkeit',
        'DE 2024 (Referenz)': 'Personenbezogene Daten von Kunden und Mitarbeitern sind unter strikter Beachtung der DSGVO und lokaler Datenschutzgesetze zu schützen.',
        'DE 2026 (Referenz)': 'Personenbezogene Daten sowie vertrauliche Kundendaten sind unter strikter Beachtung der DSGVO zu schützen; die Eingabe vertraulicher Firmendaten in öffentliche Generative-KI-Modelle ist untersagt.',
        'EN 2024': 'Personal data of customers and employees must be protected in strict compliance with the GDPR and local data protection laws.',
        'CS 2024': 'Osobní údaje zákazníků a zaměstnanců musí být chráněny v přísném souladu s GDPR a místními zákony o ochraně osobních údajů.',
        'FR 2024': 'Les données personnelles des clients et des collaborateurs doivent être protégées dans le strict respect du RGPD et des lois locales sur la protection des données.',
        'ES 2024': 'Los datos personales de clientes y empleados deben protegerse en estricto cumplimiento del RGPD y las leyes locales de protección de datos.',
        'IT 2024': 'I dati personali di clienti e dipendenti devono essere protetti nel rigoroso rispetto del GDPR e delle leggi locali sulla protezione dei dati.',
        'PL 2024': 'Dane osobowe klientów i pracowników muszą być chronione z zachowaniem ścisłej zgodności z RODO i lokalnymi przepisami o ochronie danych.',
        'NL 2024': 'Persoonsgegevens van klanten en medewerkers moeten worden beschermd in strikte overeenstemming met de AVG en lokale privacywetgeving.',
        'SK 2024': 'Osobné údaje zákazníkov a zamestnancov sa musia chrániť v prísnom súlade s GDPR a miestnymi zákonmi o ochrane osobných údajov.',
        'HU 2024': 'Az ügyfelek és a munkavállalók személyes adatait a GDPR és a helyi adatvédelmi törvények szigorú betartásával kell védeni.',
        'RO 2024': 'Datele cu caracter personal ale clienților și angajaților trebuie protejate în conformitate strictă cu GDPR și legile locale privind protecția datelor.'
      },
      {
        'Artikel-Nr.': 'Art. 08',
        'Bereich': 'Künstliche Intelligenz (Neu 2026)',
        'DE 2024 (Referenz)': '',
        'DE 2026 (Referenz)': 'Der betriebliche Einsatz von Systemen Künstlicher Intelligenz muss transparent, nachvollziehbar und unter Beachtung der geltenden KI-Verordnung (EU AI Act) sowie des Urheberrechts erfolgen.',
        'EN 2024': '',
        'CS 2024': '',
        'FR 2024': '',
        'ES 2024': '',
        'IT 2024': '',
        'PL 2024': '',
        'NL 2024': '',
        'SK 2024': '',
        'HU 2024': '',
        'RO 2024': ''
      },
      {
        'Artikel-Nr.': 'Art. 09',
        'Bereich': 'Lieferkettensorgfalt (Neu 2026)',
        'DE 2024 (Referenz)': '',
        'DE 2026 (Referenz)': 'Unsere Lieferanten und Geschäftspartner sind verpflichtet, ökologische Mindeststandards einzuhalten und faire Löhne ohne Zwangsarbeit oder Kinderarbeit nachweislich zu garantieren.',
        'EN 2024': '',
        'CS 2024': '',
        'FR 2024': '',
        'ES 2024': '',
        'IT 2024': '',
        'PL 2024': '',
        'NL 2024': '',
        'SK 2024': '',
        'HU 2024': '',
        'RO 2024': ''
      },
      {
        'Artikel-Nr.': 'Art. 10',
        'Bereich': 'Hygieneregeln & Pandemie (Entfallen)',
        'DE 2024 (Referenz)': 'Aufgrund behördlicher Auflagen gelten im Betrieb gesonderte Pandemie-Schutzmaßnahmen, Mindestabstände von 1,5 Metern sowie die Pflicht zur täglichen Erfassung der Anwesenheit.',
        'DE 2026 (Referenz)': '',
        'EN 2024': 'Due to regulatory requirements, special pandemic protective measures, a minimum distance of 1.5 meters, and mandatory daily attendance recording apply on company premises.',
        'CS 2024': 'Na základě úředních nařízení platí v provozu mimořádná protipandemická opatření, minimální rozestupy 1,5 metru a povinnost denní evidence docházky.',
        'FR 2024': 'En raison des exigences réglementaires, des mesures sanitaires spécifiques, une distance minimale de 1,5 mètre et l\'enregistrement quotidien des présences s\'appliquent.',
        'ES 2024': 'Debido a los requisitos normativos, se aplican en las instalaciones medidas especiales de protección contra pandemias, distancia mínima de 1,5 metros y registro diario de asistencia.',
        'IT 2024': 'A causa di disposizioni normative, nei locali aziendali si applicano speciali misure di protezione pandemica, distanze minime di 1,5 metri e registrazione giornaliera delle presenze.',
        'PL 2024': 'Z uwagi na wymogi urzędowe w zakładzie obowiązują szczególne środki ochrony pandemicznej, minimalne odstępy 1,5 metra oraz obowiązek codziennej rejestracji obecności.',
        'NL 2024': 'Vanwege wettelijke voorschriften gelden op de bedrijfslocatie speciale beschermingsmaatregelen, een minimumafstand van 1,5 meter en verplichte dagelijkse aanwezigheidsregistratie.',
        'SK 2024': 'Na základe úradných nariadení platia v prevádzke mimoriadne protipandemické opatrenia, minimálne odstupy 1,5 metra a povinnosť dennej evidencie dochádzky.',
        'HU 2024': 'A hatósági előírások miatt a telephelyen különleges járványvédelmi intézkedések, 1,5 méteres távolságtartás és kötelező napi jelenléti ív van érvényben.',
        'RO 2024': 'Din cauza cerințelor de reglementare, se aplică măsuri speciale de protecție împotriva pandemiei, distanța minimă de 1,5 metri și înregistrarea zilnică a prezenței.'
      },
      {
        'Artikel-Nr.': 'Art. 11',
        'Bereich': 'Fax-Meldungen (Entfallen)',
        'DE 2024 (Referenz)': 'Dringende vertrauliche Meldungen an den Compliance-Beauftragten können im Ausnahmefall auch per Telefax an die zentrale Leitstelle übermittelt werden.',
        'DE 2026 (Referenz)': '',
        'EN 2024': 'Urgent confidential reports to the compliance officer may, in exceptional cases, also be transmitted by fax to the central control station.',
        'CS 2024': 'Naléhavá důvěrná hlášení pracovníkovi compliance lze ve výjimečných případech zasílat také telefaxem na centrální dispečink.',
        'FR 2024': 'Les signalements urgents et confidentiels au responsable de la conformité peuvent, à titre exceptionnel, être transmis par télécopie au centre de contrôle.',
        'ES 2024': 'Los informes confidenciales urgentes al responsable de cumplimiento podrán remitirse, excepcionalmente, por telefax a la central.',
        'IT 2024': 'Le segnalazioni confidenziali urgenti al responsabile della conformità possono essere inviate, in casi eccezionali, anche tramite fax al centro di controllo.',
        'PL 2024': 'Pilne zgłoszenia poufne do pełnomocnika ds. zgodności można w wyjątkowych przypadkach przesyłać także faksem do centrali.',
        'NL 2024': 'Dringende vertrouwelijke meldingen aan de compliance officer kunnen bij uitzondering ook per fax aan de centrale meldpost worden verstuurd.',
        'SK 2024': 'Naliehavé dôverné hlásenia pracovníkovi compliance je možné vo výnimočných prípadoch zasielať aj faxom na centrálu.',
        'HU 2024': 'A megfelelőségi felelősnek címzett sürgős bizalmas bejelentések kivételes esetben faxon is elküldhetők a központba.',
        'RO 2024': 'Notificările confidențiale urgente adresate responsabilului cu conformitatea pot fi transmise, în cazuri excepționale, prin fax la dispeceratul central.'
      },
      {
        'Artikel-Nr.': 'Art. 12',
        'Bereich': 'Klimaschutz & Umwelt',
        'DE 2024 (Referenz)': 'Wir verpflichten uns zur kontinuierlichen Reduzierung unserer Treibhausgasemissionen und streben Klimaneutralität im Geschäftsbetrieb bis zum Jahr 2040 an.',
        'DE 2026 (Referenz)': 'Wir verpflichten uns zur kontinuierlichen Reduzierung unserer Treibhausgasemissionen und streben Klimaneutralität im gesamten Geschäftsbetrieb bis zum Jahr 2035 an.',
        'EN 2024': 'We commit to continuously reducing our greenhouse gas emissions and aim for climate neutrality in business operations by 2040.',
        'CS 2024': 'Zavazujeme se k neustálému snižování emisí skleníkových plynů a usilujeme o dosažení klimatické neutrality v provozu do roku 2040.',
        'FR 2024': 'Nous nous engageons à réduire continuellement nos émissions de gaz à effet de serre et visons la neutralité climatique des opérations d\'ici 2040.',
        'ES 2024': 'Nos comprometemos a reducir continuamente nuestras emisiones de gases de efecto invernadero y aspiramos a la neutralidad climática operativa para 2040.',
        'IT 2024': 'Ci impegniamo a ridurre costantemente le nostre emissioni di gas serra e puntiamo alla neutralità climatica nelle attività operative entro il 2040.',
        'PL 2024': 'Zobowiązujemy się do ciągłej redukcji emisji gazów cieplarnianych i dążymy do neutralności klimatycznej w działalności operacyjnej do 2040 roku.',
        'NL 2024': 'Wij verbinden ons ertoe onze uitstoot van broeikasgassen voortdurend te verminderen en streven naar klimaatneutraliteit in onze bedrijfsvoering tegen 2040.',
        'SK 2024': 'Zaväzujeme sa k neustálemu znižovaniu emisií skleníkových plynov a usilujeme sa o dosiahnutie klimatickej neutrality v prevádzke do roku 2040.',
        'HU 2024': 'Elkötelezettek vagyunk az üvegházhatású gázok kibocsátásának folyamatos csökkentése mellett, és 2040-re klímasemlegességet célzunk meg működésünkben.',
        'RO 2024': 'Ne angajăm să reducem continuu emisiile de gaze cu efect de seră și vizăm neutralitatea climatică în operațiunile companiei până în anul 2040.'
      },
      {
        'Artikel-Nr.': 'Art. 13',
        'Bereich': 'Whistleblowing & Meldungen',
        'DE 2024 (Referenz)': 'Mitarbeiter können vermutete Verstöße schriftlich an den Ombudsmann melden; eine erste Rückmeldung erfolgt innerhalb von 14 Kalendertagen.',
        'DE 2026 (Referenz)': 'Mitarbeiter können vermutete Verstöße vertraulich oder anonym über das digitale Hinweisgeberportal melden; eine erste Rückmeldung erfolgt innerhalb von 7 Werktagen.',
        'EN 2024': 'Employees may report suspected violations in writing to the ombudsman; initial feedback will be provided within 14 calendar days.',
        'CS 2024': 'Zaměstnanci mohou podezření na porušení pravidel nahlásit písemně ombudsmanovi; první zpětná vazba bude poskytnuta do 14 kalendářních dnů.',
        'FR 2024': 'Les collaborateurs peuvent signaler par écrit les manquements présumés au médiateur ; un premier retour sera effectué dans un délai de 14 jours calendaires.',
        'ES 2024': 'Los empleados pueden comunicar sospechas de infracción por escrito al defensor del empleado; la primera respuesta se facilitará en un plazo de 14 días naturales.',
        'IT 2024': 'I dipendenti possono segnalare presunte violazioni per iscritto all\'ombudsman; un primo riscontro sarà fornito entro 14 giorni di calendario.',
        'PL 2024': 'Pracownicy mogą zgłaszać podejrzenia naruszeń na piśmie do rzecznika praw; pierwsza informacja zwrotna zostanie przekazana w ciągu 14 dni kalendarzowych.',
        'NL 2024': 'Werknemers kunnen vermoedelijke overtredingen schriftelijk melden bij de ombudsman; binnen 14 kalenderdagen volgt een eerste reactie.',
        'SK 2024': 'Zamestnanci môžu podozrenia na porušenie pravidiel nahlásiť písomne ombudsmanovi; prvá spätná väzba bude poskytnutá do 14 kalendárnych dní.',
        'HU 2024': 'A munkavállalók a gyanított szabályszegéseket írásban jelenthetik az ombudsmannak; az első visszajelzés 14 naptári napon belül megtörténik.',
        'RO 2024': 'Angajații pot raporta în scris presupusele încălcări către ombudsman; un prim răspuns va fi furnizat în termen de 14 zile calendaristice.'
      },
      {
        'Artikel-Nr.': 'Art. 14',
        'Bereich': 'Schutz vor Repressalien',
        'DE 2024 (Referenz)': 'Hinweisgeber, die in gutem Glauben handeln, sind vollumfänglich vor Kündigung, Benachteiligung oder sonstigen Repressalien geschützt.',
        'DE 2026 (Referenz)': 'Hinweisgeber, die in gutem Glauben handeln, sind vollumfänglich vor Kündigung, Benachteiligung oder sonstigen Repressalien geschützt.',
        'EN 2024': 'Whistleblowers acting in good faith are fully protected against termination, discrimination, or any other retaliatory measures.',
        'CS 2024': 'Oznamovatelé jednající v dobré víře jsou plně chráněni před výpovědí, znevýhodněním či jakýmikoli jinými odvetnými opatřeními.',
        'FR 2024': 'Les lanceurs d\'alerte agissant de bonne foi sont intégralement protégés contre tout licenciement, désavantage ou mesure de représailles.',
        'ES 2024': 'Los denunciantes que actúen de buena fe están plenamente protegidos contra despidos, perjuicios o cualquier otra represalia.',
        'IT 2024': 'Gli informatori che agiscono in buona fede sono pienamente protetti da licenziamenti, discriminazioni o qualsiasi altra ritorsione.',
        'PL 2024': 'Sygnaliści działający w dobrej wierze podlegają pełnej ochronie przed zwolnieniem, dyskryminacją lub wszelkimi innymi represjami.',
        'NL 2024': 'Klokkenluiders die te goeder trouw handelen, zijn volledig beschermd tegen ontslag, benadeling of andere vergeldingsmaatregelen.',
        'SK 2024': 'Oznamovatelia konajúci v dobrej viere sú plne chránení pred výpoveďou, znevýhodnením či akýmikoľvek inými odvetnými opatreniami.',
        'HU 2024': 'A jóhiszeműen eljáró bejelentők teljes körű védelmet élveznek a felmondással, hátrányos megkülönböztetéssel vagy megtorlással szemben.',
        'RO 2024': 'Avertizorii de integritate care acționează cu bună-credință sunt pe deplin protejați împotriva concedierii, discriminării sau altor represalii.'
      },
      {
        'Artikel-Nr.': 'Art. 15',
        'Bereich': 'Schlussbestimmungen',
        'DE 2024 (Referenz)': 'Verstöße gegen diesen Kodex können arbeitsrechtliche Konsequenzen bis hin zur fristlosen Kündigung nach sich ziehen.',
        'DE 2026 (Referenz)': 'Verstöße gegen diesen Kodex können arbeitsrechtliche Konsequenzen bis hin zur fristlosen Kündigung sowie zivil- und strafrechtliche Schritte nach sich ziehen.',
        'EN 2024': 'Violations of this Code may result in employment disciplinary action up to and including termination without notice.',
        'CS 2024': 'Porušení tohoto kodexu může mít za následek pracovněprávní důsledky až po okamžité zrušení pracovního poměru.',
        'FR 2024': 'Toute violation du présent code est susceptible d\'entraîner des sanctions disciplinaires pouvant aller jusqu\'au licenciement immédiat.',
        'ES 2024': 'Las infracciones de este Código pueden dar lugar a consecuencias laborales, incluido el despido disciplinario sin preaviso.',
        'IT 2024': 'Le violazioni di questo Codice possono comportare conseguenze disciplinari fino al licenziamento per giusta causa senza preavviso.',
        'PL 2024': 'Naruszenia niniejszego Kodeksu mogą skutkować konsekwencjami dyscyplinarnymi, łącznie z natychmiastowym rozwiązaniem umowy o pracę.',
        'NL 2024': 'Overtredingen van deze code kunnen leiden tot arbeidsrechtelijke gevolgen, tot en met ontslag op staande voet.',
        'SK 2024': 'Porušenie tohto kódexu môže mať za následok pracovnoprávne dôsledky až po okamžité skončenie pracovného pomeru.',
        'HU 2024': 'A jelen Kódex megsértése munkajogi következményeket, akár azonnali hatályú felmondást is vonhat maga után.',
        'RO 2024': 'Încălcarea prezentului Cod poate atrage consecințe disciplinare, mergând până la desfacerea imediată a contractului de muncă.'
      }
    ];
  }

  function getSampleHeaders() {
    return [
      'Artikel-Nr.',
      'Bereich',
      'DE 2024 (Referenz)',
      'DE 2026 (Referenz)',
      'EN 2024',
      'CS 2024',
      'FR 2024',
      'ES 2024',
      'IT 2024',
      'PL 2024',
      'NL 2024',
      'SK 2024',
      'HU 2024',
      'RO 2024'
    ];
  }

  function getSampleDefaultMappings() {
    return {
      refOldCol: 'DE 2024 (Referenz)',
      refNewCol: 'DE 2026 (Referenz)',
      targetCols: [
        'CS 2024',
        'EN 2024',
        'FR 2024',
        'ES 2024',
        'IT 2024',
        'PL 2024',
        'NL 2024',
        'SK 2024',
        'HU 2024',
        'RO 2024'
      ]
    };
  }

  return {
    getSampleData: getSampleData,
    getSampleHeaders: getSampleHeaders,
    getSampleDefaultMappings: getSampleDefaultMappings
  };
}));
