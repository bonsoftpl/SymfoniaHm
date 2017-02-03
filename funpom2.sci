//"funpom2.sci","funpom2.sci",12001,0,1.0.5,SYSTEM

//-------------------------------------------------------------------------------
int sub Assert(int bWyr, string sTresc)
//-------------------------------------------------------------------------------
  Assert = 1
  #ifndef NO_ASSERTS
  if !bWyr then message "B³¹d wewnêtrzny raportu - niespe³niona asercja:\n" + sTresc : Assert = 0
  #endif
endsub

// xml {{{

//----------------------------------------------------------------------------
string sub XmlGetText(Dispatch node, string sPath) //{{{
//----------------------------------------------------------------------------
  Dispatch found
  found = node.selectSingleNode(sPath)
  if found then XmlGetText = found.text
  node = fp_nothing : found = fp_nothing
endsub //}}}

//----------------------------------------------------------------------------
long sub XmlGetInt(Dispatch node, string sPath) //{{{
//----------------------------------------------------------------------------
  Dispatch found
  found = node.selectSingleNode(sPath)
  if found then XmlGetInt = Val(found.text)
  node = fp_nothing : found = fp_nothing
endsub //}}}

//}}} xml

//----------------------------------------------------------------------------
string sub PierwszaSeria_fp(int bxt, string sKod) //{{{
//----------------------------------------------------------------------------
  long idDefDok

  SetKey(bxt, "kod")
  SetKeySeg(bxt, "typi", 32)
  SetKeySeg(bxt, "kod", sKod)
  if GetRec(bxt, GE) == 0 && GetKeySeg(bxt, "typi") == 32 && GetKeySeg(bxt, "kod") == sKod then
    idDefDok = GetField(bxt, "id")
    SetKey(bxt, "super")
    SetKeySeg(bxt, "super", idDefDok)
    GetRec(bxt, GE)
    while BaseError(bxt, 0) == 0 && GetKeySeg(bxt, "super") == idDefDok
      if !!(GetField(bxt, "flag") & 128) then
        PierwszaSeria_fp = GetField(bxt, "kod")
        exit
      endif
      GetRec(bxt, NX)
    wend
  endif
endsub //}}}

//---------------------------------------------------------------
string sub Eskapuj(string s) //{{{
//---------------------------------------------------------------
  string sWyn, sLit
  long i, L, nAscii
  
  L = Len(s)
  for i=1 to i>L
    sLit = Mid(s, i, 1)
    nAscii = sLit(1)
    if nAscii < 0 then nAscii += 256
    if nAscii < 32 then
      sWyn += (using "&#%03l;", nAscii)
    else
      sWyn += sLit
    endif
  next i
  Eskapuj = sWyn
endsub //}}}

//---------------------------------------------------------------
string sub OdEskapuj(string s) //{{{
//---------------------------------------------------------------
  string sTmp
  int nPoz, nAscii

  nPoz = move 0 : sTmp = buf : buf = s
  
  while find regular "/&/#{[0-9][0-9][0-9]};"
    nAscii = Val(regular 1)
    delete 6
    insert (using "%c", nAscii)
  wend
  OdEskapuj = buf

  buf = sTmp : move to 1 : move nPoz-1
endsub //}}}

//----------------------------------------------------------------------------
float sub AbsF(float f) //{{{
//----------------------------------------------------------------------------
  if f >= 0 then
    AbsF = f
  else
    AbsF = -f
  endif
endsub //}}}

//----------------------------------------------------------------------------
float sub MinF(float f1, float f2) //{{{
//----------------------------------------------------------------------------
  if f1 < f2 then
    MinF = f1
  else
    MinF = f2
  endif
endsub //}}}

//----------------------------------------------------------------------------
int sub AppendLineToFile(string sFile, string sLine) //{{{
//----------------------------------------------------------------------------
  Dispatch fso
  Dispatch f
  fso = CreateObject("Scripting.FileSystemObject")
  f = fso.OpenTextFile (sFile, 8, -1)
  f.Write(sLine + "\r\n")
  f.Close
  fso = fp_nothing : f = fp_nothing
endsub //}}}

//----------------------------------------------------------------------------
int sub LogReportRun(string sRepName) //{{{
//----------------------------------------------------------------------------
  AppendLineToFile(Katalog() + "bs_rap_log.txt", sRepName + CurrentUser() + ", " + Data() + ", " + Time())
endsub //}}}
