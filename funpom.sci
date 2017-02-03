//"funpom.sci","funpom.sci",12001,0,1.0.7,SYSTEM

// wersja 1.0.15

#define FUNPOM
#define BS_FUNPOM
#define WM_CLOSE 0x0010
int anDniWMies(12) = 31,28,31,30,31,30,31,31,30,31,30,31
int bExcelKropka = 0
string sZera="000000000000000000000000000000000000"

#define idbasXT   2
#define idbasPR   6
#define idbasKH  20
#define idbasBD  28
#define idbasDK  34

Dispatch fp_nothing

//--------------------------------------------------------
int sub DumpMap (int fDeb, mapvalue m)
//--------------------------------------------------------
  int i
  if fDeb<=0 then exit
  for i=1 to i>m.Size()
    print #fDeb; (using "Mapa %3d: %s", i, m.GetKey(i))
    print #fDeb; " "; m.Get(i); lf
  next i
endsub

//--------------------------------------------------------
string sub MapSafeGetS (mapvalue m, string sKey)
//--------------------------------------------------------
  long im
  im = m.Index(sKey)
  if im then MapSafeGetS = m.Get(im)
endsub

//--------------------------------------------------------
float sub MapSafeGetF (mapvalue m, string sKey)
//--------------------------------------------------------
  long im
  im = m.Index(sKey)
  if im then MapSafeGetF = m.Get(im)
endsub

//--------------------------------------------------------
long sub MapSafeGetN (mapvalue m, string sKey)
//--------------------------------------------------------
  long im
  im = m.Index(sKey)
  if im then MapSafeGetN = m.Get(im)
endsub

//------------------------------------------------------------------
string sub MEditString (string s, string sTitle, string sText)
//------------------------------------------------------------------
  string sPom = s
  int wys,y
  Form sTitle, 300,wys=220
    Text sText, 10,y=10, 100,21
    MEdit "", sPom, 10,y+22, 270,100
    Button "&OK", 60,wys-70, 70,30, 2
    Button "&Anuluj", 170,wys-70, 70,30, -1
  if (ExecForm)==2 then
    MEditString = sPom
  else
    MEditString = s
  endif
endsub

//------------------------------------------------------------------
string sub StringToMEdit (string s)
//------------------------------------------------------------------
  // w polu MEdit entery s¹ podane jako \r
  string sTmp
  long nPoz
  
  nPoz = (move 0) : sTmp = buf
  
  buf = s
  while replace "\r", "" : wend
  while find "\n" : delete "\n" : insert "\r\n" : move 2 : wend
  StringToMedit = buf
  
  buf = sTmp : move to nPoz
endsub

//-------------------------------------------------------------------------------
string sub KatalogZeSciezki(string sSciezka)
//-------------------------------------------------------------------------------
  string sTmp
  sTmp = buf
  buf = sSciezka
  if find regular at "^{*}\\[~\\]+$" then
    KatalogZeSciezki = regular 1
  else
    KatalogZeSciezki = sSciezka
  endif
  buf = sTmp
endsub

//-------------------------------------------------------------------------------
string sub PlikZeSciezki(string sSciezka)
//-------------------------------------------------------------------------------
  string sTmp
  sTmp = buf
  buf = sSciezka
  if find regular "\\{[~\\]++}$" then
    PlikZeSciezki = regular 1
  else
    PlikZeSciezki = sSciezka
  endif
  buf = sTmp
endsub

//------------------------------------------------------------------------------
int sub StartsWith(string s1, string s2)
//------------------------------------------------------------------------------
  if Mid(s1, 1, Len(s2))==s2 then StartsWith = 1
endsub

//------------------------------------------------------------------------------
int sub EndsWith(string s1, string s2)
//------------------------------------------------------------------------------
  int L1, L2
  L1 = Len(s1)
  L2 = Len(s2)
  if L2 <= L1 && Mid(s1, L1+1-L2) == s2 then EndsWith = 1
endsub

//------------------------------------------------------------------------------
string sub DodajStrSep(string s1, string s2, string sSep)
//------------------------------------------------------------------------------
  DodajStrSep = s1
  if s2 then
    if s1 then DodajStrSep += sSep
    DodajStrSep += s2
  endif
endsub

//----------------------------------------------------------------------------
int sub ListaPrzecDoMapy(string sLista, mapvalue m) //{{{
//----------------------------------------------------------------------------
  string sTmp
  int nPoz

  nPoz = move 0 : sTmp = buf : buf = sLista
  m.Clear()
  while find regular at ":b{[~,]+}:b(,|$)-"
    m.Set((regular 1), 1)
  wend
  if (move 0)<=Len(buf) then message "ListaPrzecDoMapy: b³¹d wewnêtrzny raportu, skontaktuj siê z autorem."
  
  buf = sTmp : move to 1 : move nPoz-1
endsub //}}}

//-------------------------------------------------------------------------------
int sub PlikIstnieje(string sPlik)
//-------------------------------------------------------------------------------
  int h = open sPlik for input
  if h>0 then PlikIstnieje=1 : close h
endsub


//------------------------------------------------------------------
string sub InputBS (string s, string sTitle, string sText)
//------------------------------------------------------------------
  string sPom = s
  int wys,y
  Form sTitle, 300,wys=140
    Text sText, 10,y=10, 100,21
    Edit "", sPom, 10,y+22, 270,21
    Button "&OK", 60,wys-70, 70,30, 2
    Button "&Anuluj", 170,wys-70, 70,30, -1
  if (ExecForm)==2 then
    InputBS = sPom
  else
    InputBS = s
  endif
endsub

//---------------------------------------------------------
string sub DataPlusNDni (string sData, long n)
//---------------------------------------------------------
  Date d
  d.FromStr(sData)
  if d.Valid() then
    d.Add(n)
    DataPlusNDni = d.ToStr()
  endif
endsub

//---------------------------------------------------------
long sub DataMinusData (string sData2, string sData1)
//---------------------------------------------------------
  long wyn
  int r1=Val(Mid(sData1,1,4)), m1=Val(Mid(sData1,6,2)), d1=Val(Mid(sData1,9,2))
  int r2=Val(Mid(sData2,1,4)), m2=Val(Mid(sData2,6,2)), d2=Val(Mid(sData2,9,2))
  int nDniWMies

  wyn+=d2-1 : d2=1 : wyn-=d1-1 : d1=1

  while (r2-r1)*12+(m2-m1)>0
    nDniWMies = anDniWMies(m1)
    if m1==2 then
      // luty, trzeba sprawdziæ przestêpnoœæ roku
      if r1%4==0 && (r1%100!=0 || r1%400==0) then nDniWMies += 1
    endif
    m1+=1
    if m1>12 then m1-=12 : r1+=1
    wyn += nDniWMies
  wend

  DataMinusData = wyn
endsub

//------------------------------------------------ GetRecById {{{
int sub GetRecById(int b, long id)
//---------------------------------------------------------------
  SetKey(b,"id")
  SetKeySeg(b,"id",id)
  GetRecById = GetRec(b,EQ)
endsub //}}}

//---------------------------------------------------------------
int sub GetRecByKod(int b, string sKod, string sTyp)
//---------------------------------------------------------------
  SetKey(b,"kod")
  SetKeySeg(b,"typ",sTyp)
  SetKeySeg(b,"typi",Val(sTyp))
  SetKeySeg(b,"kod",sKod)
  GetRecByKod = GetRec(b,EQ)
endsub

//---------------------------------------------------------------
int sub GetRecBySuperKod(int b, long nSuper, string sKod)
//---------------------------------------------------------------
  SetKey(b, "super")
  SetKeySeg(b, "super", nSuper)
  SetKeySeg(b, "kod", sKod)
  GetRecBySuperKod = GetRec(b, EQ)
endsub

//---------------------------------------------------------------
long sub PodajIdKat (int bxt, string sKod, string sTyp, long lSuper)
//---------------------------------------------------------------
  SetKey(bxt,"kod")
  SetKeySeg(bxt,"typ",sTyp)
  SetKeySeg(bxt,"kod",sKod)
  if GetRec(bxt,GE)==0 && GetKeySeg(bxt,"typ")==sTyp && GetKeySeg(bxt,"kod")==sKod then
    PodajIdKat = GetField(bxt,"id")
  else
    PodajIdKat = lSuper
  endif
endsub

//----------------------------------------------------------------------------
string sub PelnaNazwaKat(int bxt, long id) //{{{
//----------------------------------------------------------------------------
  string s

  if !id || id == 1 then exit
  if GetRecById(bxt, id) != 0 then exit
  s = "\\@" + GetField(bxt, "kod")
  s = PelnaNazwaKat(bxt, GetField(bxt, "super")) + s
  PelnaNazwaKat = s
endsub //}}}

//----------------------------------------------------------------------------
string sub PelnaNazwaKatObr(int bxt, long id) //{{{
//----------------------------------------------------------------------------
  // obrobiona do postaci akceptowanej przez iorec
  long i, L, iPocz, cSeg
  long cPominSeg
  string s
 
  cPominSeg = 1
  s = PelnaNazwaKat(bxt, id)
  L = Len(s)

  for i=1 to i>L
    if s(i) == '\\' then cSeg += 1
  next i
  //if cSeg == cPominSeg then cPominSeg -= 1

  iPocz = 1
  for i=2 to i>L
    if s(i) == '\\' && cPominSeg > 0 then
      cPominSeg -= 1
      iPocz = i
    endif
  next i

  PelnaNazwaKatObr = Mid(s, iPocz)
endsub

//---------------------------------------------------------------
long sub DajIdKontr(int b, string sKod)
//---------------------------------------------------------------
  if !sKod then exit
  SetKey(b,"kod")
  SetKeySeg(b,"typ","0")
  SetKeySeg(b,"kod",sKod)
  if GetRec(b,EQ)==0 then DajIdKontr=GetField(b,"id")
endsub

//---------------------------------------------------------------
string sub DajKodKontr(int b, long id)
//---------------------------------------------------------------
  if !id then exit
  SetKey(b,"id")
  SetKeySeg(b,"id",id)
  if GetRec(b,EQ)==0 then
    DajKodKontr=GetField(b,"kod")
  else
    message "B³¹d odczytu z bazy kontrahentów."
  endif
endsub

//----------------------------------------------------------------------------
string sub DajZnaczStr(int nZnacz) //{{{
//----------------------------------------------------------------------------
  // nZnacz - GetField(xx, "znaczniki")
  if nZnacz then
    if nZnacz>='A' && nZnacz<='Z' then
      DajZnaczStr = (using "%c", nZnacz)
    else
      DajZnaczStr = (using "%l", nZnacz - 'Z' - 1)
    endif
  endif
endsub //}}}

//---------------------------------------------------------------
string sub DajZnaczKontr(int b, string sKod)
//---------------------------------------------------------------
  if !sKod then exit
  SetKey(b,"kod")
  SetKeySeg(b,"typ","0")
  SetKeySeg(b,"kod",sKod)
  if GetRec(b,EQ)==0 then DajZnaczKontr=GetField(b,"znacznik")
endsub

//--------------------------------------------------------------
long sub DataToLong (string sData)
//--------------------------------------------------------------
  if sData then
    DataToLong = 0x00010000*Val(Mid(sData,1,4)) + 0x00000100*Val(Mid(sData,6,2)) + Val(Mid(sData,9,2))
  else
    DataToLong = 1
  endif
endsub

//--------------------------------------------------------------
string sub LongToData (long lData)
//--------------------------------------------------------------
  if lData>1 then
    LongToData = (using "%04l-%02l-%02l", (lData&0xFFFF0000)/0x00010000, (lData&0x0000FF00)/0x00000100, lData&0x000000FF)
  endif
endsub

//--------------------------------------------------------------
string sub GetLongName(int bGetLongNames, int b, int bnt, string sPole)
//--------------------------------------------------------------
  int bJestDluga
  if bGetLongNames && GetField(b,"idlongname") then
    SetKey(bnt,"id") : SetKeySeg(bnt,"id",GetField(b,"idlongname"))
    GetRec(bnt,EQ)
    if BaseError(bnt,2)!=0 then
      message "B³¹d przy odczycie d³ugiej nazwy dla:\n"+GetField(b,sPole)
    else
      GetLongName = GetField(bnt,"opis")
      bJestDluga = 1
    endif
  endif

  if !bJestDluga then
    GetLongName = GetField(b,sPole)
  endif
endsub

//--------------------------------------------------------------------
int sub SetLongName (string sNotka, int b, int bnt, string sTyp, int nBaza, long lSuper)
//--------------------------------------------------------------------
  if sNotka then
    if GetField(b,"idlongname") then
      if GetRecById (bnt, GetField(b,"idlongname"))!=0 then message "B³¹d odczytu z bazy notatek." : exit
      SetField(bnt,"opis",sNotka)
      PutRec(bnt) : if BaseError(bnt,2)!=0 then AbortTrans():close:error""
    else
      Clear(bnt)
      SetField(bnt,"opis",sNotka)
      SetField(bnt,"typ", sTyp)
      SetField(bnt,"typi", Val(sTyp))
      SetField(bnt,"baza", nBaza)
      SetField(bnt,"super", lSuper)
      InsRec(bnt) : if BaseError(bnt,2)!=0 then AbortTrans():close:error""
      SetField(b, "idlongname", GetField(bnt,"id"))
      PutRec(b) : if BaseError(b,2)!=0 then AbortTrans():close:error""
    endif
  else
    // nie ma notki
    if GetField(b,"idlongname") then
      if GetRecById (bnt, GetField(b,"idlongname"))!=0 then message "B³¹d odczytu z bazy notatek." : exit
      DelRec (bnt) : BaseError(bnt,4)
      SetField (b, "idlongname", 0)
      PutRec (b) : BaseError(b,4)
    endif
  endif
endsub

//----------------------------------------------------------------------------
string sub DajNotkeRek(int bnt, string sBaza, int nTyp, long id) //{{{
//----------------------------------------------------------------------------
  // bazê mo¿na podaæ numerem lub nazw¹
  // nTyp w typowych sytuacjach jest równy 0
  // trzeba mieæ ustawiony limit 8192
  if sBaza == "DK" then sBaza = "16"
  if sBaza == "BD" then sBaza = "28"
  if sBaza == "BM" then sBaza = "38"
  SetKey(bnt, "super")
  SetKeySeg(bnt, "typi", nTyp)
  SetKeySeg(bnt, "baza", Val(sBaza))
  SetKeySeg(bnt, "super", id)
  SetKeySeg(bnt, "typ", (using "%l", nTyp))
  if GetRec(bnt, EQ) == 0 then DajNotkeRek = GetField(bnt, "opis")
endsub //}}}

//---------------------------------------------------------------
string sub NazwaTowaru(long idtw, int btw, int bnt)
//---------------------------------------------------------------
  if !idtw then exit
  if GetRecById(btw, idtw)==0 then
    NazwaTowaru = GetLongName(1, btw, bnt, "nazwa")
  else
    NazwaTowaru = (using "#B£¥D ODCZYTU idtw=%l", idtw)
  endif
endsub

//---------------------------------------------------------------
int sub MessageYesNo (string sMessage, string sTitle, int bDefTak)
//---------------------------------------------------------------
  int y, cLinie, wys
  string sTmp
  int nPoz
  string sPom

  if !sTitle then sTitle = "Potwierdzenie"
  nPoz = move 0 : sTmp = buf
  buf = sMessage

  // najpierw policzymy ile linii jest w komunikacie, ¿eby ewentualnie wyd³u¿yæ okienko
  cLinie = 1
  while find regular "(\n\r)|(\r\n)|(\n)-"
    cLinie += 1
  wend
  move to 0

  wys=200
  if cLinie>3 then wys += ((cLinie-3)*35)

  y=-20
  Form sTitle, 300,wys
    while find regular at "^*(\n\r)|(\r\n)|(\n)|($)-"
      sPom = delete from begin
      Text sPom, 20,y+=35, 270,45
    wend
    if bDefTak then
      Button "&Tak", 60,wys-70, 70,30, 2
      Button "&Nie", 170,wys-70, 70,30, -1
    else
      Button "&Nie", 170,wys-70, 70,30, -1
      Button "&Tak", 60,wys-70, 70,30, 2
    endif
  MessageYesNo = (ExecForm == 2)
  buf = sTmp : move to 1 : move nPoz-1
endsub

//---------------------------------------------------------------
int sub SubSave (int rv)
//---------------------------------------------------------------
  Save
  SubSave = rv
endsub

//---------------------------------------------------------------
int sub UtworzPrawoUzytk (int bpr, string sPrawo)
//---------------------------------------------------------------
  long idKatal
  
  // sprawdzimy, czy pod id 65528 figuruje skrot " Inne "
  // je¿eli tak (ver>=300f), to wsadzamy prawa pod to pole
  // w przeciwnym wypdaku pod 65501
  if GetRecById(bpr,65528)==0 && Mid(GetField(bpr,"skrot"),2,4)=="Inne" then
    idKatal=65528
  else
    idKatal=65501
  endif

  Clear(bpr)
  SetField(bpr,"katalog",idKatal)
  SetField(bpr,"skrot",sPrawo)
  if InsRec(bpr)!=0 then
    message "Niemo¿liwe dodanie prawa:\n"+sPrawo
    if BaseError(bpr,0)!=5 then BaseError(bpr,2)
  else
    UtworzPrawoUzytk = 1
  endif
endsub

//------------------------------------------------------------------------
int sub UtworzGalazPrawUzytk (int bpr, string sGalaz)
//------------------------------------------------------------------------
  Clear(bpr)
  SetField(bpr,"flag",14)
  SetField(bpr,"subtyp","62")
  SetField(bpr,"typ",16)
  SetField(bpr,"katalog",65501)
  SetField(bpr,"skrot",sGalaz)
  if InsRec(bpr)!=0 then
    message "Nie uda³o siê utworzyæ ga³êzi praw u¿ytkownika:\n"+sGalaz
    if BaseError(bpr,0)!=5 then BaseError(bpr,2)
  endif
endsub

//---------------------------------------------------------------
int sub UtworzPrawoUzytkWGal (int bpr, string sPrawo, string sGalaz)
//---------------------------------------------------------------
  // tworzy prawo u¿ytkownika w ga³êzi (od wersji 3.00f) o podanej nazwie
  long idKatal
  SetKey(bpr,"skrot")
  SetKeySeg(bpr,"skrot",sGalaz)
  if GetRec(bpr,EQ)==0 then
    idKatal = GetField(bpr,"id")
  else
    idKatal = 65528  // standardowa ga³¹Ÿ "Inne"
  endif
  
  Clear(bpr)
  SetField(bpr,"katalog",idKatal)
  SetField(bpr,"skrot",sPrawo)
  if InsRec(bpr)!=0 then
    message "Niemo¿liwe dodanie prawa:\n"+sPrawo
    if BaseError(bpr,0)!=5 then BaseError(bpr,2)
  endif
endsub

// wersja 3.42 wprowadzi³a funkcjê o tej samej nazwie
// funkcja standardowa pobiera tylko jeden parametr: id prawa
//---------------------------------------------------------------
int sub CzyMaPrawoJC (int bpr, int bkh, int bzz, string sUser, string sPrawo)
//---------------------------------------------------------------
  long idUser, idPrawo
  if !sUser then sUser = CurrentUser
  if sUser=="Admin" then CzyMaPrawoJC = 1 : exit
  SetKey(bkh,"kod")
  SetKeySeg(bkh,"typ","103")
  SetKeySeg(bkh,"kod",sUser)
  GetRec(bkh,EQ) : BaseError(bkh,4)
  idUser = GetField(bkh,"id")

  SetKey(bpr,"skrot")
  SetKeySeg(bpr,"skrot",sPrawo)
  if GetRec(bpr,EQ)!=0 then exit
  idPrawo = GetField(bpr,"id")

  SetKey(bzz,"cross1")
  SetKeySeg(bzz,"typ","32")
  SetKeySeg(bzz,"baza1",20) : SetKeySeg(bzz,"id1",idUser)
  SetKeySeg(bzz,"baza2",6) : SetKeySeg(bzz,"id2",idPrawo)
  if GetRec(bzz,EQ)!=0 then exit
  CzyMaPrawoJC = (GetField(bzz,"cena") != 0)
endsub

//---------------------------------------------------------------
string sub ImieINazwiskoUzytkownika (int bkh, string sUser)
//---------------------------------------------------------------
  GetRecByKod (bkh, sUser, "103")
  ImieINazwiskoUzytkownika = GetField(bkh,"nazwa")
endsub

//----------------------------------------------------------------------
long sub PodajMagUzytk (int bxt, int bzz, string sUser)
//----------------------------------------------------------------------
  long idUser
  SetKey(bxt,"super")
  SetKeySeg(bxt,"super",6000)
  SetKeySeg(bxt,"kod",sUser)
  if GetRec(bxt,EQ)!=0 then close : error "Nie mo¿na odczytaæ danych u¿ytkownika "+sUser+" z bazy XT."
  idUser = GetField(bxt,"id")

  SetKey (bzz,"cross1")
  SetKeySeg (bzz, "typ", "102")
  SetKeySeg (bzz, "baza1", 2)
  SetKeySeg (bzz, "id1", idUser)
  SetKeySeg (bzz, "baza2", 2)
  GetRec(bzz,GE)
  if BaseError(bzz,0)!=0 || GetKeySeg(bzz,"typ")!="102" || GetKeySeg(bzz,"baza1")!=2 || GetKeySeg(bzz,"id1")!=idUser || GetKeySeg(bzz,"baza2")!=2 then close : error "Nie mo¿na odczytaæ magazynu u¿ytkownika "+sUser+" z bazy ZZ."
  PodajMagUzytk = GetField (bzz, "id2")
endsub

#ifndef _BTR_SCI

//--------------------------------------------
int sub openWC (string sPlik, string sBaza)
//--------------------------------------------
  int f = open sPlik for base sBaza
  if BaseError(sPlik,0)==0 then openWC=f : exit
  buf = sPlik
  if find regular "-\\[~\\]++$" then delete to end
  f = open buf+"\\con" for input
  if f then
    close(f)
  else
    if 0 != MkDir(buf) then message "Nie mo¿na utworzyæ katalogu\n"+buf : close:error""
  endif
  int b = open sPlik for base sBaza
  if BaseError(sPlik,0)==12 then
    // pliku nie da siê otworzyæ, mo¿na siê by³o spodziewaæ
    Create sPlik For Base sBaza
    if BaseError(sPlik,0)!=0 then
      message (using "Nie mo¿na utworzyæ pliku:\n%s\ndla bazy %s", sPlik, sBaza) : close:error""
    endif
    b = open sPlik for base sBaza
  endif
  openWC = b
endsub

#endif

//---------------------------------------------------------------------
string sub LiczbaExcel (float fKw, int nMiejsca)
//---------------------------------------------------------------------
  string sUs
  if nMiejsca then
    sUs = (using "%%.%df", nMiejsca)
  else
    sUs = "%.f"
  endif
  if bExcelKropka then
    LiczbaExcel = (using sUs, fKw)
  else
    buf = (using sUs, fKw)
    replace ".", ","
    LiczbaExcel = buf
  endif
endsub

//---------------------------------------------------------------------
string sub StawkaVATzID (int bxt, int subtyp)
//---------------------------------------------------------------------
  // to dzia³a trochê powoli, bo trzeba jeszcze usuwaæ znaki ascii o kodach œmieciach
  int ch, i
  string sWynik

  SetKey ( bxt,"super" )
  SetKeySeg( bxt,"super",10000 )
  GetRec (bxt,GE)
  while BaseError(bxt,0)==0 && GetKeySeg(bxt,"super")==10000
    if GetField(bxt,"subtyp") == (using "%d",subtyp) then
      sWynik = GetField(bxt,"kod")
      // wytniemy jeszcze znaki o kodzie >127
      for i=1 to i>len(sWynik)
        ch = sWynik(i)
        if ch<0 then ch += 256
        if ch<128 then StawkaVATzID += Mid(sWynik,i,1)
      next i
      exit
    endif
    GetRec(bxt,NX)
  wend
endsub

//----------------------------------------------------------------------------
int sub ZnajdzRejRekurs(int bxt, long nNumerRej, long idSuper) //{{{
//----------------------------------------------------------------------------
  // to jest funkcja prywatna
  // zwraca 1 gdy znajdzie
  mapvalue midChildren
    midChildren.Type(int)
  int bMam
  long im

  SetKey(bxt, "super")
  SetKeySeg(bxt, "super", idSuper)
  GetRec(bxt, GE)
  while BaseError(bxt,0)==0 && GetKeySeg(bxt,"super")==idSuper
    midChildren.Set((using "%l", GetField(bxt, "id")), 1)
    if GetField(bxt, "long") == nNumerRej then
      bMam = 1
      exit
    endif
    GetRec(bxt, NX)
  wend

  if bMam then
    ZnajdzRejRekurs = 1
  else
    for im = 1 to im > midChildren.Size()
      if ZnajdzRejRekurs(bxt, nNumerRej, Val(midChildren.GetKey(im))) then
        bMam = 1
        ZnajdzRejRekurs = 1
        exit
      endif
    next im
  endif
endsub

//----------------------------------------------------------------------------
string sub ZnajdzRej(int bxt, long nNumerRej) //{{{
//----------------------------------------------------------------------------
  // numer rejestru jest w polu rejestrVat(i) dokumentu
  // oraz w xt w polu long
  // je¿eli znajdzie odpowiedni rejestr, zwraca jego nazwê i ustawia rekord xt
  if ZnajdzRejRekurs(bxt, nNumerRej, 10400) then
    ZnajdzRej = GetField(bxt, "nazwa")
  endif
endsub //}}}

//-------------------------------------------------------------------------------
string sub FmtIlosc (float fIlosc, int nMiejscaMin, int nMiejscaMax)
//-------------------------------------------------------------------------------
  string sTmp
  int iKropka, nZera
  
  sTmp = buf
  if nMiejscaMax>=0 then
  	// obcinamy dalsze miejsca, je¿eli podano nMiejscaMax
    buf = (using (using "%%.%df", nMiejscaMax), fIlosc)
  else
    buf = (using "%f", fIlosc)
  endif

  iKropka = 0
  if find "." then iKropka = move 0
  if !iKropka then buf += "." : iKropka=len(buf)

  // zjadamy koñcowe zera + ewentualnie kropkê
  while len(buf) && Mid(buf,len(buf),1)=="0"
    buf = Mid(buf,1,len(buf)-1)
  wend
  
  // dope³niamy zerami, je¿eli podano nMiejscaMin
  if nMiejscaMin then
    nZera = nMiejscaMin - (len(buf) - iKropka)
    if nZera>0 then buf += Mid(sZera,1,nZera)
  endif
  
  // je¿eli na koñcu wyniku otrzymaliœmy kropkê, to wywalmy j¹
  if len(buf) && Mid(buf,len(buf),1)=="." then buf = Mid(buf,1,len(buf)-1)

  #ifdef FMTILOSC_PRZEC
    move to 0
    replace ".", ","
  #endif

  FmtIlosc = buf
  buf = sTmp
endsub

//------------------------------------------------------------------
float sub BS_Abs(float f)
//------------------------------------------------------------------
  if f<0 then
    BS_Abs = -f
  else
    BS_Abs = f
  endif
endsub

//------------------------------------------------------------------
int sub MessageBufArg()
//------------------------------------------------------------------
  string s
  s = (using "buf=%s;arg0=%s;arg1=%s;arg2=%s;arg3=%s;arg4=%s;arg5=%s", buf, arg0, arg1, arg2, arg3, arg4, arg5)
  s += (using "\narg6=%s;arg7=%s;arg8=%s;arg9=%s", arg6, arg7, arg8, arg9)
  message s
endsub

//-------------------------------------------------------------------------------
int sub DumpIoRec(int hDeb, Iorec ior)
//-------------------------------------------------------------------------------
  int bJestSekcja
  string sOds
  string sSpacje = "                                                             "
  
  if hDeb<=0 then exit
  bJestSekcja = 1
  ior.SetAtFirst()
  while bJestSekcja
    sOds = Mid(sSpacje, 1, 2*ior.GetLevel())
    print #hDeb; sOds; ior.GetSectionName(); " {"; lf
    while ior.NextField()
      print #hDeb; sOds; "  ";ior.GetFieldName(); "="; ior.GetFieldValue(); lf
    wend
    bJestSekcja = ior.NextSection()
  wend
endsub

//-------------------------------------------------------------------------------
int sub StoreArgs()
//-------------------------------------------------------------------------------
  //message "StoreArgs"
  PutIni("BONSOFT", "arg0", arg0)
  PutIni("BONSOFT", "arg1", arg1)
  PutIni("BONSOFT", "arg2", arg2)
  PutIni("BONSOFT", "arg3", arg3)
  PutIni("BONSOFT", "arg4", arg4)
  PutIni("BONSOFT", "arg5", arg5)
  PutIni("BONSOFT", "arg6", arg6)
  PutIni("BONSOFT", "arg7", arg7)
  PutIni("BONSOFT", "arg8", arg8)
  PutIni("BONSOFT", "arg9", arg9)
endsub

//-------------------------------------------------------------------------------
int sub RestoreArgs()
//-------------------------------------------------------------------------------
  //message "RestoreArgs"
  arg0 = GetIni("BONSOFT", "arg0")
  arg1 = GetIni("BONSOFT", "arg1")
  arg2 = GetIni("BONSOFT", "arg2")
  arg3 = GetIni("BONSOFT", "arg3")
  arg4 = GetIni("BONSOFT", "arg4")
  arg5 = GetIni("BONSOFT", "arg5")
  arg6 = GetIni("BONSOFT", "arg6")
  arg7 = GetIni("BONSOFT", "arg7")
  arg8 = GetIni("BONSOFT", "arg8")
  arg9 = GetIni("BONSOFT", "arg9")
endsub

//-------------------------------------------------------------------------------
int sub DialogFolder(string sPrompt, long idEB)
//-------------------------------------------------------------------------------
  string sKat
  Dispatch sh
  sh.Create("Shell.Application")
  sKat = GetVal(idEB)
  Dispatch rv = sh.BrowseForFolder(0, sPrompt, 1)
  if rv then
    sKat = rv.ParentFolder.ParseName(rv.Title).Path
    SetVal(idEB, sKat)
  endif
endsub

//-------------------------------------------------------------------------------
int sub WyswietlPomoc(string sSkrotRap)
//-------------------------------------------------------------------------------
  // uwaga - wymaga wysokiego limit
  int bpr, h
  string sPlikTmp

  buf = Katalog()+"amhm51pr.dat"
  #ifdef FORTE
  buf = ""
  #endif
  bpr = open buf for base "PR"
  if BaseError(buf,2)!=0 then message "B³¹d podczas otwierania bazy PR" : exit
  SetKey(bpr, "skrot")
  SetKeySeg(bpr, "skrot", sSkrotRap)
  if GetRec(bpr,EQ)==0 then
    buf = GetField(bpr,"dane")
    delete regular "^////[~\n\r]+[\n\r]##"
    sPlikTmp = GetTempFileName("bs_", "htm")
    h = open sPlikTmp for output
    if h>0 then
      print #h;buf
      close h
      ShellExecute(sPlikTmp, "open")
      message "Pomoc otwar³a siê w oknie przegl¹darki internetowej"
      delete file sPlikTmp
    else
      message "B³¹d podczas zapisu do pliku tymczasowego"
    endif
  else
    message "B³¹d odczytu pliku raportu "+sSkrotRap
  endif
  close bpr

endsub

//----------------------------------------------------------------------------
int sub WyswietlPomocPdf(string sProjekt) //{{{
//----------------------------------------------------------------------------
  string sPlik
  int h
  sPlik = Katalog() + "dokumentacja\\" + sProjekt + ".pdf"
  h = open sPlik for binary input
  if h <= 0 then message "Brak pliku " + sPlik : exit
  close h
  ShellExecute(sPlik, "open")
endsub //}}}

// DajMnem //{{{
/// m - mapa wykorzystanych mnemoników, zaktualizowana
/// zwraca tekst sOpcja z wstawionym ampersandem
//----------------------------------------------------------------------------
string sub DajMnem(string sOpcja, mapvalue m)
//----------------------------------------------------------------------------
  long L, i
  string sLit

  DajMnem = sOpcja
  L = Len(sOpcja)
  for i=1 to i>L
    sLit = Ucase(Mid(sOpcja, i, 1))
    if !m.Index(sLit) then
      m.Set(sLit, 1)
      DajMnem = Mid(sOpcja, 1, i-1) + "&" + Mid(sOpcja,i)
      exit
    endif
  next i
endsub //}}}

//*********************** FUNKCJE WALIDUJ¥CE ***********************
int sub DataOk (string sData)
  buf = sData
  if (find regular "^[1-9][0-9]^3/-(0[1-9])|(1[012])/-(0[1-9])|([12][0-9])|(3[01])$")!="" then DataOk=1
endsub

#include "funpom2.sci"
