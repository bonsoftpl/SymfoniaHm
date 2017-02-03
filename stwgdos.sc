//"stwgdos.sc","Stany wed³ug dostaw",12303,0,2.0.2,SYSTEM

#include "funpom.sci"

//******************** PARAMETRY I STA£E ***************
#define MIN_STAN 0.0000001
//#define DIAG
limit 10000
#ifdef DIAG
  int hOut = open "c:\\temp\\1\\debug.txt" for output
#endif

//************************** TYPY DANYCH ********************
record trDane
  string sKod[25+1]
  string sNazwa[100+1]
  float fCena
  float fIlosc
  float fWart
  string sStVat[5+1]
  string sJm[8+1]
  string sKodDost[25+1]
  long idtw
  int sub Zeruj()
    sKod="" : sNazwa="" : fCena=0 : fIlosc=0 : fWart=0 : sStVat=""
    sJm="" : sKodDost="" : idtw=0
  endsub
endrec

//**************************** ZMIENNE GLOBALNE *************************************
string sIniName="Stany magazynowe wed³ug dostaw"
int cMag
//int bStanHandl
int bBezZerowych, bNaDzien, bNazwaTowaru, bPodzialNaVat, bBezPodsumTow, bExcel
float fPom
long y
string sDataRap, sMag, sFiltrDost, sPom
long g_idtw
//int bPods
long idMag
Dispatch nothing
Dispatch e
Dispatch wb
dispatch sh
mapvalue msMag
  msMag.Type(string)
mapvalue mnMag
  mnMag.Type(int)
mapvalue midMag
  midMag.Type(long)
trDane arDane(1)
mapvalue miDane
  miDane.Type(int)
string sPoleStan
mapvalue msStawkiVat
  msStawkiVat.Type(string)
mapvalue mfStawkiVat
  mfStawkiVat.Type(float)  // stopa stawki Vat, np. 0.22 dla 22%


// TABELE I STYLE
int cKol=6, tab, tabn, tabp
int st_tl,st_tc,st_tr, st_gl
string asNagl(cKol) = "Towar", "Dostawa", "Stan", "Jm.", "Wartoœæ", "Cena"
int ad(cKol) =            500,       500,    200,   150,       200,    200
string asst(cKol)=       "tl",       "tl",   "tr", "tc",      "tr",   "tr"
long aw(cKol) =             1,         1,      0,     0,         0,      0

int odsp, kpods0=1, kpods1=3, kpods2=cKol
int cKolp = kpods2-kpods1+1+1
float afSuma(cKol), afSumaTow(cKol), afSumaVat(cKol)

mapvalue mStyl
  mStyl.Type(int)

#include "Formatowanie wydruków"

int bxt = open KatalogFirmy()+"51xt.dat" for base "XT" : BaseError(KatalogFirmy()+"51xt.dat",4)
int bsm = open KatalogFirmy()+"51sm.dat" for base "SM" : BaseError(KatalogFirmy()+"51sm.dat",4)
int btw = open KatalogFirmy()+"51tw.dat" for base "TW" : BaseError(KatalogFirmy()+"51tw.dat",4)
int bdw = open KatalogFirmy()+"51dw.dat" for base "DW" : BaseError(KatalogFirmy()+"51dw.dat",4)
int bpw = open KatalogFirmy()+"51pw.dat" for base "PW" : BaseError(KatalogFirmy()+"51pw.dat",4)
int bnt = open KatalogFirmy()+"51nt.dat" for base "NT" : BaseError(KatalogFirmy()+"51nt.dat",4)

int idEB,idCB
int sub OnComm (int id, int msg)
  Enable (idEB, -1+2*GetVal(idCB))
endsub
//-------------------------------------------------------------------------------
int sub Dialog()
//-------------------------------------------------------------------------------
  int rv, szer,wys, x,y, i
  string asMag(1)
  
  for i=1 to i>msMag.Size()
    asMag(Size(asMag)) = msMag.Get(i)
    grow asMag, 1
  next i
  asMag(Size(asMag))="*** wszystkie ***"
  
  bExcel = Val(GetIni (sIniName, "excel"))
  bBezZerowych = Val(GetIni (sIniName, "Bez zerowych"))
  bNaDzien = Val(GetIni (sIniName, "Na dzien"))
  sDataRap = GetIni (sIniName, "Data raportu")
  bNazwaTowaru = Val(GetIni (sIniName, "Nazwa towaru"))
  bPodzialNaVat = Val(GetIni (sIniName, "Podzial na VAT"))
  bBezPodsumTow = Val(GetIni (sIniName, "Bez podsumowan towarow"))
  sMag = GetIni (sIniName, "Magazyn")
  sFiltrDost = GetIni (sIniName, "Filtr dostaw")
  
  //bStanHandl = Val(GetIni (sIniName, "Stan handlowy"))
  //bPods = Val(GetIni (sIniName, "Podsumowanie"))

  Form sIniName, szer=300, wys=320
    Text "&Magazyn:", 20,(y=20)+2, 70,21
    CmbBox "", asMag, sMag, 90,y, 120,200
    ChkBox " Twórz arkusz E&xcel", bExcel, 20,y+=30, 150,21
    ChkBox " &Bez dostaw zerowych", bBezZerowych, 20,y+=22, 150,21
    ChkBox " B&ez podsumowañ dla towarów", bBezPodsumTow, 20,y+=22, 180,21
    ChkBox " Na&zwa towaru", bNazwaTowaru, 20,y+=22, 150,21
    ChkBox " &Podzia³ na stawki VAT", bPodzialNaVat, 20,y+=22, 150,21
    idCB=ChkBox " &Na konkretny dzieñ: ", bNaDzien, 20,y+=22, 130,21
    idEB=DatEdit "", sDataRap, 150,y, 90,21
    Text "&Filtr dostaw:", 20,(y+=30)+2, 100,21
    Edit "", sFiltrDost, 120,y, szer-120-30,21
    //ChkBox " &Podsumowanie", bPods, 20,y+=25, 150,21
    //RadioBtn "Stany &faktyczne", bStanHandl, 40,y+=30, 120,21
    //RadioBtn "Stany &handlowe", bStanHandl, 160,y, 120,21
    Button "&OK", (szer-140)/3,wys-70, 70,30, 2
    Button "&Anuluj", szer-70-(szer-140)/3,wys-70, 70,30, -1
  rv = ExecForm OnComm
  if rv<2 then close:error""
  //if bStanHandl then
  //  sPoleStan = "stanHandl"
  //else
  //  sPoleStan = "stan"
  //endif

  PutIni (sIniName, "excel", (using "%d",bExcel))
  PutIni (sIniName, "Bez zerowych", (using "%d",bBezZerowych))
  PutIni (sIniName, "Na dzien", (using "%d",bNaDzien))
  PutIni (sIniName, "Data raportu", sDataRap)
  PutIni (sIniName, "Nazwa towaru", (using "%d",bNazwaTowaru))
  PutIni (sIniName, "Podzial na VAT", (using "%d",bPodzialNaVat))
  PutIni (sIniName, "Bez podsumowan towarow", (using "%d",bBezPodsumTow))
  //PutIni (sIniName, "Stan handlowy", (using "%d",bStanHandl))
  //PutIni (sIniName, "Podsumowanie", (using "%d",bPods))
  PutIni (sIniName, "Magazyn", sMag)
  PutIni (sIniName, "Filtr dostaw", sFiltrDost)

  if sMag=="*** wszystkie ***" then sMag=""
  if sMag then
    idMag=midMag.Get(sMag)
  else
    idMag=0
  endif
endsub

//-------------------------------------------------------------------------------
int sub WczytajMagazyny()
//-------------------------------------------------------------------------------
  int ind
  SetKey(bxt,"super")
  SetKeySeg(bxt,"super",6900)
  GetRec(bxt,GE)
  while BaseError(bxt,0)==0 && GetKeySeg(bxt,"super")==6900
    msMag.Set ( (using "%l",GetField(bxt,"id")), GetField(bxt,"kod"))
    ind = mnMag.Size()+1
    mnMag.Set ( (using "%l",GetField(bxt,"id")), ind )
    midMag.Set ( GetField(bxt,"kod"), GetField(bxt,"id"))
    GetRec(bxt,NX)
  wend
  cMag=msMag.Size()
endsub

//-------------------------------------------------------------------------------
int sub DodefiniujTabele()
//-------------------------------------------------------------------------------
  int i, st
  int ftGruby = CopyFont ("tekst", 1, TextHeight("X","tekst")+5)
  st_tl=Styl("tekst",-1,"tl") : st_tc=Styl("tekst",0,"tc") : st_tr=Styl("tekst",1,"tr")
  st_gl=Styl(ftGruby, -1, "gl")
  mStyl.Set("tl",st_tl) : mStyl.Set("tc",st_tc) : mStyl.Set("tr",st_tr)
  mStyl.Set("gl",st_gl)

  long wall, dall
  int ast(cKol), astn(cKol)
  for i=1 to i>cKol
    wall += aw(i)
    dall += ad(i)
    astn(i) = st_tc
    ast(i) = mStyl.Get(asst(i))
  next i

  for i=1 to i>cKol
    ad(i) += aw(i)*(str.szer-dall)/wall
  next i

  // tabela podsumowania
  int adp(cKolp), astp(cKolp)
  int k=1
  for i=1 to i>cKol
    if i<kpods0 then odsp += ad(i)
    if i>=kpods0 && i<kpods1 then adp(k)+=ad(i)
    if i>=kpods1 && i<=kpods2 then adp(k+=1) = ad(i) : astp(k)=mStyl.Get(asst(i))
  next i
  astp(1) = st_tc

  tab = tabela 2,10, ad,ast
  tabn = tabela 2,10, ad,astn
  tabp = tabela 2,10, adp,astp
endsub


//-------------------------------------------------------------------------------
int sub NaglowekVat (string sStawka)
//-------------------------------------------------------------------------------
  SetStyl("gl")
  print "Stawka VAT "; sStawka; lf
  print at #X,#Y+5;
endsub
    
//-------------------------------------------------------------------------------
int sub StopkaVat (string sStawka)
//-------------------------------------------------------------------------------
  int i,k
  tabela #tabp, od odsp,#Y+20
    kolumna k=1, "Razem stawka "+sStawka
    kolumna k+=1, FmtIlosc(afSumaVat(k-2+kpods1), -1, -1)
    kolumna k+=2, Kwota(afSumaVat(k-2+kpods1))
  koniec
  print at #X,#Y+20;
  for i=1 to i>cKol
    afSumaVat(i) = 0.00
  next i
endsub

//-------------------------------------------------------------------------------
float sub StanTowaruSM(long idtw)
//-------------------------------------------------------------------------------
  SetKey(bsm,"towar")
  SetKeySeg(bsm, "idtw", idtw)
  SetKeySeg(bsm, "magazyn", idMag)
  if GetRec(bsm,EQ)==0 then
    StanTowaruSM = GetField(bsm,"wartosc")
  else
    StanTowaruSM = 0
  endif
endsub

//-------------------------------------------------------------------------------
int sub NaglowekTabeli()
//-------------------------------------------------------------------------------
  int k
  tabela #tabn
    for k=1 to k>cKol
      kolumna k, asNagl(k)
    next k
  koniec
endsub
    
//-------------------------------------------------------------------------------
int sub StopkaTabeli (int bOstatnia)
//-------------------------------------------------------------------------------
  float fStanSM, fStanDW
  int i,k

  if bOstatnia then
    tabela #tabp, od odsp,#Y+20
      kolumna k=1, "----- PODSUMOWANIE -----"
      kolumna k+=1, FmtIlosc(afSuma(k-2+kpods1), -1, -1)
      kolumna k+=2, Kwota(afSuma(k-2+kpods1))
    koniec
    if bExcel then
      y += 1
      sh.Cells(y,cKol-3).FormulaR1C1 = "=SUM(R[-1]C:R2C)"
      sh.Cells(y,cKol-1).FormulaR1C1 = "=SUM(R[-1]C:R2C)"
      sh.Range(sh.Cells(2,cKol), sh.Cells(y,cKol)).NumberFormat = "0,00"
    endif
  else

#ifdef DIAG
  if !g_idtw then message "Nie okreœlone g_idtw" : close:error""
  fStanSM = StanTowaruSM(g_idtw)
  fStanDW = afSumaTow(kpods1+2)
  if Sign(fStanSM-fStanDW,2)!=0 then
    GetRecById(btw, g_idtw)
    if hOut>0 then print #hOut; GetField(btw,"kod"), fStanSM, fStanDW; lf
  endif
#endif

    if !bBezPodsumTow then
      tabela #tabp, od odsp,#Y
        kolumna k=1, "Razem towar"
        kolumna k+=1, FmtIlosc(afSumaTow(k-2+kpods1), -1, -1)
        kolumna k+=2, Kwota(afSumaTow(k-2+kpods1))
      koniec
      print at #X,#Y+20;
    endif

    for i=1 to i>cKol
      afSuma(i) += afSumaTow(i)
      afSumaVat(i) += afSumaTow(i)
      afSumaTow(i) = 0
    next i

  endif

endsub

int sub Opis(int bPierwszy)
  NaglowekTabeli()
endsub

int sub Podsumowanie(int bOstatnie)
endsub

//-------------------------------------------------------------------------------
int sub Inicjalizacja()
//-------------------------------------------------------------------------------
  long k
	if bExcel then
		Popup(1, "Uruchamiamy Excela")
		e.Create("Excel.Application")
		e.ScreenUpdating = 0
    sh = e.Workbooks.Add.Sheets(1)
    y = 1
		
		//if bDebug then
		//  e.Visible = -1
		//  e.ScreenUpdating = -1
		//endif
		
		for k=1 to k>ckol
		  sh.Cells(y,k).Formula = "'" + asNagl(k)
		next k
	endif

endsub

//-------------------------------------------------------------------------------
int sub CloseAll()
//-------------------------------------------------------------------------------
  int i
  sh=nothing : wb=nothing : e=nothing
  close
endsub

//-------------------------------------------------------------------------------
int sub CloseErr()
//-------------------------------------------------------------------------------
  if e then e.visible=1
  CloseAll : error ""
endsub

//-------------------------------------------------------------------------------
int sub WczytajStawkiVat()
//-------------------------------------------------------------------------------
  int ch, i
  string sStawka, sStawka2

  SetKey ( bxt,"super" )
  SetKeySeg( bxt,"super",10000 )
  GetRec (bxt,GE)
  while BaseError(bxt,0)==0 && GetKeySeg(bxt,"super")==10000
    //if GetField(bxt,"subtyp") == (using "%d",subtyp) then
    sStawka = GetField(bxt,"kod")
    sStawka2 = ""
    // wytniemy jeszcze znaki o kodzie >127
    for i=1 to i>len(sStawka)
      ch = sStawka(i)
      if ch<0 then ch += 256
      if ch<128 then sStawka2 += Mid(sStawka,i,1)
    next i
    msStawkiVat.Set( GetField(bxt,"subtyp"), sStawka2 )
    mfStawkiVat.Set( GetField(bxt,"subtyp"), GetField(bxt,"wartosc") )
    GetRec(bxt,NX)
  wend
endsub

//-------------------------------------------------------------------------------
int sub ZaktualizujDostaweNaDzien (int bdw, string sDataRap)
//-------------------------------------------------------------------------------
  // aktualizujemy pola stan, wartoscst tak, by by³y prawid³owe na wybrany dzieñ
  // w momencie odczytu rekordu pola odzwierciedlaj¹ stan aktualny
  // po wykonaniu SetField dalsza czêœæ raportu przebiega bez zmian
  long iddw = GetField(bdw,"id")
  float fStan, fWart
  
  // stanu pocz¹tkowego nie bêdziemy nawet pobieraæ z DW, tylko z PW
  
  SetKey (bpw, "pozycje")
  SetKeySeg (bpw, "typ", "37")
  SetKeySeg (bpw, "iddw", iddw)
  GetRec(bpw,GE)
  while BaseError(bpw,0)==0 && GetKeySeg(bpw,"iddw")==iddw && GetKeySeg(bpw,"typ")=="37"
  	if GetField(bpw,"data")<=sDataRap then
  	  fStan -= GetField(bpw,"ilosc")
  	  fWart -= GetField(bpw,"wartosc")
  	endif
    GetRec(bpw,NX)
  wend

  SetField (bdw, "stan", fStan)
  //SetField (bdw, "wartoscst", GetField(bdw,"stan")*GetField(bdw,"cena"))
  SetField (bdw, "wartoscst", fWart)
endsub

//-------------------------------------------------------------
int sub DrukujDane()
//-------------------------------------------------------------
  int nKey, i,j,k, bPierwszy
  trDane r
  string sStVat, sStVatOld
  long idtwOld

  miDane.Sort()
  bPierwszy = 1

  for nKey=1 to nKey>miDane.Size()
    r = arDane(miDane.Get(nKey))

    if idtwOld!=r.idtw then
      if afSumaTow(1)>0.0 then    // je¿eli wypuszczono jakiœ rekord
        g_idtw = idtwOld
        StopkaTabeli(0)
      endif
    endif
    idtwOld = r.idtw

    sStVat = r.sStVat
    if bPodzialNaVat && (bPierwszy || sStVat!=sStVatOld) then
      if !bPierwszy then StopkaVat(sStVatOld)
      NaglowekVat(sStVat)
    endif
    bPierwszy = 0
    sStVatOld = sStVat
    if bExcel then y+=1

    tabela #tab
      kolumna k=1
        if bNazwaTowaru then
          print sPom=r.sNazwa
        else
          print sPom=r.sKod
        endif
        if bExcel then sh.Cells(y,k).Formula = "'" + sPom
      kolumna k+=1, sPom=r.sKodDost
        if bExcel then sh.Cells(y,k).Formula = "'" + sPom
      kolumna k+=1
        if r.fIlosc>MIN_STAN || r.fIlosc<-MIN_STAN then
          print FmtIlosc(fPom = r.fIlosc, -1, -1)
          if bExcel then sh.Cells(y,k).Formula = fPom
        else
          if bExcel then sh.Cells(y,k).Formula = 0
        endif
        afSumaTow(k) += r.fIlosc
      kolumna k+=1, sPom=r.sJm
        if bExcel then sh.Cells(y,k).Formula = "'" + sPom
      kolumna k+=1
        if r.fWart>MIN_STAN || r.fWart<-MIN_STAN then
          print Kwota(fPom = r.fWart)
          if bExcel then sh.Cells(y,k).Formula = fPom
        else
          if bExcel then sh.Cells(y,k).Formula = 0
        endif
        afSumaTow(k) += r.fWart
      kolumna k+=1
        //if GetField(bdw,"stan")>MIN_STAN then
          print Kwota(fPom=r.fCena)
          if bExcel then sh.Cells(y,k).Formula = fPom
        //endif
      afSumaTow(1) += 1 // zliczamy po prostu liczbê wydrukowanych rekordów
    koniec

  next nKey

  if afSumaTow(1)>0.0 then    // je¿eli wypuszczono jakiœ rekord
    g_idtw = r.idtw
    StopkaTabeli(0)
  endif
  if bPodzialNaVat then StopkaVat(sStVat)
  StopkaTabeli(1)
endsub

//*****************************************************************************************
//********************************** PROGRAM G£ÓWNY ***************************************
//*****************************************************************************************

str.Wydruk(0,-1,-1)
strona 150,150,100,150
SetStyl (Styl("tekst",-1))

WczytajMagazyny()
Dialog()
Inicjalizacja()
WczytajStawkiVat()
DodefiniujTabele()
SetStyl (Styl("tekst",-1))
nazwaRaportu = sIniName
if sMag then nazwaRaportu += (", magazyn "+sMag)
if bNaDzien then
  nazwaRaportu += (" na dzieñ " + sDataRap)
else
  nazwaRaportu += (", wydruk z dnia " + Data())
endif
grow afSuma, cMag-1

int bOk, i, k, bZerowy, ind
long idTow, lp
string sField, sKey
float afStan(cMag)
trDane r

Header(1)
bOk = SetTaggedPos(FS)
lp = 0
while bOk
  idTow = GetLineId()
  bZerowy = 1

  SetKey(btw,"id") : SetKeySeg(btw,"id",idTow) : GetRec(btw,EQ)
  if BaseError(btw,0)!=0 then message (using "B³¹d odczytu z bazy towary (id=%l)",idTow) : BaseError(btw,2)

  if GetField(btw,"subtyp")!="0" then goto NastTowar

  for i=1 to i>cKol : afSumaTow(i)=0 : next i
  Popup (1, GetField(btw,"kod"))

  SetKey(bdw,"towar")
  SetKeySeg(bdw,"idtw", GetField(btw,"id"))
  SetKeySeg(bdw,"numer", 0)   // numery mog¹ byæ te¿ ujemne, ale je pomijamy
  GetRec(bdw,GE)

  while BaseError(bdw,0)==0 && GetKeySeg(bdw,"idtw")==GetField(btw,"id")
    if idMag && GetField(bdw,"magazyn")!=idMag then goto nastDost
    if sFiltrDost then
      buf = GetField(bdw,"kod")
      if !find regular at "^=-" + sFiltrDost + "$" then goto nastDost
    endif
    if bNaDzien then
      if GetField(bdw,"data")>sDataRap then goto nastDost
      ZaktualizujDostaweNaDzien (bdw, sDataRap)
    endif
    if !bBezZerowych || GetField(bdw,"stan")>MIN_STAN || GetField(bdw,"stan")<-MIN_STAN then
      // tu musimy zebraæ dane do tabelki
      r.Zeruj()
      r.sKod = GetField(btw,"kod")  //+(using "  %s %s", GetField(btw,"typ"), GetField(btw,"subtyp"))
      r.sNazwa = GetLongName (1, btw, bnt, "nazwa")
      r.sKodDost = GetField(bdw,"kod")
      r.fIlosc = GetField(bdw,"stan")
      r.sJm = GetField(btw,"jm")
      r.fWart = Round(GetField(bdw,"wartoscst"), 2)
      r.fCena = GetField(bdw,"cena")
      r.sStVat = MapSafeGetS(msStawkiVat, (GetField(btw,"vatsp")) )
      r.idtw = GetField(btw,"id")
      //r.sNazwa += ("_"+r.sStVat)
      
      if bPodzialNaVat then
        sKey = (using "%s%c%08l", r.sStVat, 255, lp+=1)
      else
        sKey = (using "%08l", lp+=1)
      endif
      ind = Size(arDane)
      miDane.Set (sKey, ind)
      arDane(ind) = r
      grow arDane,1
    endif

    nastDost:
    GetRec(bdw,NX)
  wend

  NastTowar:
  bOk = SetTaggedPos(NX)
wend

DrukujDane()
Footer(1)

if bExcel then
	for k=1 to k>ckol
	  sh.Columns(k).AutoFit
	next k
  e.Visible = 1
  e.ScreenUpdating = -1
endif
Popup(1,"")
CloseAll()
