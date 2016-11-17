//"eksp_rap.sc","AAA Eksport raportów",2400,0,1.0.1,SYSTEM
//"eksp_rap.sc","AAA Eksport raportów","\System\",0,1.0.1,SYSTEM
///////////////////////////////////////
// Eksport raportów
///////////////////////////////////////

int err
string kat_wyj, nazwa_wyj, sProg, sProg1, sWersja, sFirmaCode, sPlikWyj
string sPlikWyj0
int plik_wyj
mapvalue mbPlikZap
  mbPlikZap.Type(int)
limit 60000
int hLista
str.Wydruk(0, -1, -1)
strona 100, 100, 100, 100
SetStyl(Styl("tekst", -1))

//----------------------------------------------------------------------------
string sub MutujNazwePliku(string s) //{{{
//----------------------------------------------------------------------------
  string sOldBuf
  int nOldPozBuf
  long iNr

  nOldPozBuf = move 0 : sOldBuf = buf
  buf = s
  if find regular "_{[0-9]++}/." then
    iNr = Val(regular 1)
    delete 1 + Len(regular 1)
  endif
  iNr += 1
  if find regular "/." then
    insert (using "_%03l", iNr)
  endif
  if buf == s then
    message "B³¹d wewnêtrzny Bonsoft: MutujNazwePliku zwróci³o tê sam¹ nazwê: " + s
    close : error ""
  endif
  MutujNazwePliku = buf
  
  buf = sOldBuf : move to 1 : move nOldPozBuf - 1
endsub //}}}

string sub wczytaj_string (string prompt, string war_pocz)
string wynik = war_pocz
Form prompt, 400, 110
  Edit "", wynik, 80, 10, 200, 24
  Button "&OK", 20, 50, 80, 24, 2
  Button "&Anuluj", 300, 50, 80, 24, -1
int nRetV = execform
if nRetV==0 || nRetV==-1 then close : error ""
wczytaj_string = wynik
endsub

int bprg
#ifdef HMF
  sFirmaCode = xFactory.GetObject("BProgram").kod
#else
  sFirmaCode = firma.code
#endif
select case sFirmaCode
  case "HMF"
    bprg = Open KatalogFirmy() for base "PR"
    sProg1 = "hm"
    sProg = "hmf"
  case "MKP"
    bprg = Open katalog()+"mkp61pr.dat" for base "PR"
    sProg = sProg1 = "mp"
  case "FP"
    bprg = Open katalog()+"amfp51pr.dat" for base "PR"
    sProg = sProg1 = "fp"
  case else
    bprg = Open katalog()+"amhm51pr.dat" for base "PR"
    sProg = sProg1 = "hm"
endselect
if !bprg then error "B³¹d przy otwieraniu bazy PROGRAMY !!!"
#ifdef HMF
  buf = xFactory.GetObject("BProgram").wersja
#else
  buf = firma.ver
#endif
delete regular "^20"
delete "." : delete "."
if Len(buf)==2 then buf += "0"
if Len(buf)>2 && Mid(buf,3,1)>="a" then buf = Mid(buf,1,2) + "0" + Mid(buf,3)
sWersja = buf

kat_wyj = wczytaj_string ("Podaj katalog wyjœciowy:","s:\\raporty\\std\\"+sProg1+"\\"+sProg+sWersja)
mkdir(kat_wyj)
buf = kat_wyj
if !find regular "(\\)|(/:)$" then kat_wyj += "\\"
hLista = open kat_wyj+"lista_raportow.txt" for output

SetKey (bprg, "skrot")
err = GetRec (bprg, FS)
while !err
  if sProg=="mp" then
    if GetField(bprg,"typ")!=0 then goto nastepny
  else
    if sProg != "hmf" && GetField(bprg,"id")<65536 then goto nastepny
  endif
  if GetField(bprg,"idcomp")!=0 then goto nastepny
  buf = GetField(bprg,"skrot")
  print #hLista; buf; lf
  delete regular at "?##\\"
  nazwa_wyj = buf
  buf = GetField(bprg,"dane")
  if find regular at "////\"{*}\"" then nazwa_wyj = regular 1
  if nazwa_wyj == "" then nazwa_wyj = wczytaj_string ("Podaj nazwê dla " + GetField(bprg,"skrot"),"")
  print kat_wyj+nazwa_wyj,GetField(bprg,"skrot"), GetField(bprg,"idcomp"), "..."
  if nazwa_wyj && (GetField(bprg,"idcomp")==0) then
    plik_wyj = 0
    sPlikWyj = sPlikWyj0 = kat_wyj+nazwa_wyj
    while mbPlikZap.Index(sPlikWyj)
      sPlikWyj = MutujNazwePliku(sPlikWyj)
    wend
    if sPlikWyj != sPlikWyj0 then
      print "\nKONIECZNA BY£A MUTACJA nazwy pliku: "; sPlikWyj0;" -> "; sPlikWyj; lf
    endif
    plik_wyj = open sPlikWyj for binary output
    if plik_wyj==0 then
      print "Nie uda³o siê zachowaæ pliku ";sPlikWyj;lf
    else
      print #plik_wyj; buf
      close plik_wyj
      print "Zapisany"
      mbPlikZap.Set(sPlikWyj, 1)
    endif
  endif
  print lf
  nastepny:
  err = GetRec (bprg, NX)
wend
