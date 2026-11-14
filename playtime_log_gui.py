#!/usr/bin/env python3
"""
playtime_log_gui.py
===================
Interactive GUI for playtime_log_to_excel.py (Wii Message Board dumps).

Accepts any dump form - folder, .zip, .rar, or bare .vff container - of a
"vff" dump (RECOMMENDED: holds everything cdb/cdbackup have, plus game title
IDs and all embedded photos), a "cdb" text export, or a raw container.
Two dumps of the same console can be combined; duplicates merge (the richer
copy wins, regardless of order).

Layout:
  * Left toolbar (collapsible with the ☰ button): input slots, live checklist
    of expected workbook contents, output settings, Convert button.
  * Main area: interactive, sortable views of the scanned data -
    per-title totals, per-day play time, and memos - plus a Photos tab that
    decodes attached AJPG photos (proprietary Wii format) and shows them
    alongside the message they were attached to - plus the log pane.

Requirements (Tools → Dependencies… in the app can install missing
packages for you):
  * Python 3.8+ on Windows (Tkinter is included with the python.org installer)
  * openpyxl            -> python -m pip install openpyxl   (REQUIRED)
  * Pillow              -> python -m pip install pillow     (REQUIRED for photos)
  * tkinterdnd2         -> python -m pip install tkinterdnd2   (optional, drag & drop)
  * rarfile             -> python -m pip install rarfile       (optional, .rar input;
                             also needs WinRAR/UnRAR installed for extraction)

Keep this file in the SAME FOLDER as playtime_log_to_excel.py (the converter
core). ajpg_to_png.py (the AJPG photo decoder) is also used when present,
but an embedded copy is built into this script as a fallback, so photo
decoding keeps working even when the file is missing (e.g. inside a
PyInstaller build that did not bundle it).

To build a standalone .exe (no Python needed on the target machine):
    pip install pyinstaller openpyxl pillow tkinterdnd2 rarfile
    pyinstaller --noconfirm --onefile --windowed --name PlaytimeLogConverter ^
        --collect-all tkinterdnd2 playtime_log_gui.py
"""

import base64
import contextlib
import hashlib
import importlib
import importlib.util
import io
import os
import pickle
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import types
import zipfile
import zlib
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

# --- make sure the conversion script sits next to this file -----------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
if getattr(sys, "frozen", False):  # PyInstaller build: also look next to the .exe
    sys.path.insert(0, os.path.dirname(os.path.abspath(sys.executable)))

try:
    from playtime_log_to_excel import (  # noqa: E402
        TITLE_IDS, VFF_PATH_RE, SourceBundle, build_workbook, dedup_bundles,
        is_letter, load_source, memo_category, memo_game_detail, open_source,
        parse_board_date, parse_cdbfile, parse_memo_text, readable_title,
        title_id_name, vff_record_date)
    CONVERTER_AVAILABLE = True
    CONVERTER_IMPORT_ERROR = ""
except Exception as e:  # openpyxl missing, or converter file not found
    CONVERTER_AVAILABLE = False
    CONVERTER_IMPORT_ERROR = f"{type(e).__name__}: {e}"

# Packages 1-click-installed from a packaged .exe land next to the program
# in "python_packages" (pip --target); make them importable.
_EXTRA_PKG_DIR = os.path.join(
    os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__)),
    "python_packages")
if os.path.isdir(_EXTRA_PKG_DIR) and _EXTRA_PKG_DIR not in sys.path:
    sys.path.insert(0, _EXTRA_PKG_DIR)

# --- Pillow (photo display) -------------------------------------------------
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
    PIL_IMPORT_ERROR = ""
except Exception as e:
    Image = ImageTk = None
    PIL_AVAILABLE = False
    PIL_IMPORT_ERROR = f"{type(e).__name__}: {e}"

# --- ttkbootstrap (optional themed UI with light/dark mode) ------------------
try:
    import ttkbootstrap
    TTB_AVAILABLE = True
except Exception:
    ttkbootstrap = None
    TTB_AVAILABLE = False


def system_prefers_dark():
    """Best-effort OS dark-mode detection (Windows registry, macOS defaults)."""
    try:                                   # Windows 10/11
        import winreg
        with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes"
                r"\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except Exception:
        pass
    try:                                   # macOS
        import subprocess
        out = subprocess.run(
            ["defaults", "read", "-g", "AppleInterfaceStyle"],
            capture_output=True, text=True, timeout=2)
        return "dark" in out.stdout.lower()
    except Exception:
        return False


def _bs(name):
    """ttkbootstrap colored-button style; plain ttk fallback without it."""
    return f"{name}.TButton" if TTB_AVAILABLE else "TButton"

# --- AJPG photo decoder: prefer ajpg_to_png.py, embedded copy as fallback ---
_AJPG_EMBEDDED = (
    b"c-rk;ZF3vPk^YWfF@eiDz`{^Fv+s<lI!Tr#rEEu$W&0c+k68js5-|usi-kz4uI{&AKixC43lfrZiL36eN=&?3>`c!~_tQPy)4TK8"
    b"?;dV9&BL?#@?pJvr`EUa<!TuZhQr~Y`rG>Av|XL9mlp@?Tcy?0)$(24wCbn%Tzy+_Hq}L~o~){7rXH$Sm#cQAo?ceVWxY_3|M=>;"
    b"nqN_cjiSe|zI(1#jr!xOXU`7?11Y3}%UZ3Q)w-G2ZPnbW;h(;GF%&9wzG|+jR&6)+Or70w)%(3tB~`U;HNC87gEa=QQKR$Kb~#h?"
    b"rJBypzBrgsBg1CCUYwqR<g}@$t7bM-=ko;(t)iMge*LvNI5>E?nq5Bp;vkA5HC-*+YQC(S&6q~}hPs{+)vRtX=t`Ye^Y-$5yHKx~"
    b"S8BZ?maNWsjJsSg2S$xvJpTFdm|M77H#GcJGZ;O6+&q8s_)oLT!^5uN`B_Czh?T^Kueih4&!41OI-aPeCcbZKRWC1STy;ZCId4{1"
    b"+~rg{8;}Hvi3jTOVxbnR)i2w%YO6EM@~mF0uGKY>srsgEs%cB~)Zkgw&aZq!#3$FnVDPmtf3v;1f?P%>(`UDBO%*P5y`0UfrMj#u"
    b"V#xS#py=;Gq|^{n9m>x{scmcz)$!YHwV1cJhsDWpbZAbF-)Vz?F5N%%{Nl1bG<17CpS72`lLiXcK)Y(I1<Oip=I?7MG*${D_cJ55"
    b"v;?DD4W7-^y1HGgs@Y)hiX`^!(;w8RxGB`xVm19mt?3c?y;>5Dy1uMwCP(j%%tXx(>I0P=(*FqP;OlC+xK&rwB7O~afmbgD4}{*?"
    b"c3N*pMD1l$*U}4tgjg+u`FWzAp3#2|IHlEBPY-F#FSh6BSJhItgnX{2MpxC1GKM(ZZq(0&fl_X#mn5BPtn?#hz=WqPe}Bb;zxvrq"
    b"^=m?=+Oz~~j1uZbWEV?_y<TqSt7X4|YK8$4Tp2~4DgUxj&33t*FE41YG{Jf{pSG*!fQSE>hu_rex&nqqRCA)<Rf}ysrh&|x4fO8)"
    b"{NjCeLHbS{I_&jFm5KV!oHHJ{Oz;m%ZH|s|bwvD~L8F+Pf9N$Ke=8brb3G@0Xb6RBBXb|Mt2OS_)pkQzCdAY&w@dx11!+!udC+T$"
    b"M}4R^)m7aaQ?<dYJ-5M`rDrtr5FA5q2ZJY_q-QnZeA!YhVu^#i-b|}S?PTEe@J2m;{2f@zWfS%2k+I{+0OKR{thN`IP87?{YEhFQ"
    b">KoEkg8xlD)9clo5DO%~{Oai&HTq^%pPw)4MsM2NMLjs%wr$;<FK!8IfBNqGH>XctzWM&Bm=N`nROr#8H1ttHg^GvMMRm1S-~8}~"
    b"zzaJ#72{5plpsNtN5Hf@pw%^*RdhU`&Y?guF|rsxKsiL%ZUs?yR*(5|y=@QhP#v$fEq+kR$;U{79cHT>{yxK;W*fK05UfdVPrrKd"
    b"H7vLi`@Z>Dvs#nkn_d#>Y~Cmu3;sc~jswJZ`sY^64&iA^<_<l6yv@I@>kDd9W=M@5;NQ<g{xV;TwFO`a*|Z=I2CoS>b3((RdNp6b"
    b"M32_<HP@=hhOhbAm{;P#{7P1tjlbKp&34*`U$>jVU@)uC)#*I8qZ?<JZ`ALPRCM#T`{QR#-ENyD6=^Clk^58gb9(d$wG02KUXX^}"
    b"P`jV0*VXc(){FTSv4aq!{*~*QSy`SJtV3&J)scGr_`B!NPN~B2h#GsO(y{tI)r`OX`eYEwsNOM;gg`i<OG%eF!o}cXafxwBa7l5="
    b"aLI82MGT4r6bUF2P$aQfHj{uV0aXI31XKyA5>Ta}N<o!^Dn&mjC{j?Qph!WHf+8(ln;EDwP-URXK$U?i162m93{)AYGW3&!A_qkd"
    b"iX0R<DDu>`nS&|^RSv2gR0XICP!*smKvjUMKtBa23Q!cFC_qtyqBO3}5>zFqN>G)cDnV6(ssvR@R5prmF}PS<Vq6kjQd}}za%s~>"
    b"1%83bfXaZ%fXaZ%fXaZ%fXd{!fWm;ng2IBrg2Fl(Sqmx)DhnzLDqz6^3l><gz=8!9EU;jK1q&=#V8H?l7Fe*claU1;Ebw4~2b+Kj"
    b"Sg^o?1r{u@V1We-ELdQ{0t*&cu)u-^7HsNdgqz>+@~O>FZGLL=Q=6aK{M6>BHvbQ_N$dar<o};&KGA%l`9$;oi3ZWlL{VfzaSfS1"
    b"ih?Kxq6i4FBZ?7ga=?uM86mHacoC8!$VAA901zP^0yKnN2%um_=|6;3h^i1&Q9Xoiv5Q6#ydic&=!VD*fg9pBgl&l65wIa%L%4=$"
    b"4Z#{>HH2!3i#`#oAyz}EhDZ&88sapBX)xso*b%QIT&HMVAy7x0jxarQ(Ix_Q#OVmr5v3zYM~sdT9T7SLbj0Tf&nY@r2+R?eBP`Ed"
    b"bcw(maXG?rMCAy|5tAb%M?{W*9Pv29af-$j0&&FQ2*XRqQv~9O!x4s;%zp>~QMf@AZV-hVMBxTexIq+dBFF#{xIqMNL}WhNXTh=L"
    b"SagWO4GWGX$D)Jm41{MaOBxiApMd}kBxoQ)LoytkV4vqNymgDc%Oy5JmsPuMs)gpQTt!c@$0(bAy)B@#V(^!j&%?Hkz>c><7-#}O"
    b"9(yj(;kZ6^@G=gZZcE4kE)(!DND2W8?*ZfjMoe44EISYkL4jTBAXWC@5edY2@1PVI^QJ}YAe#&PQpZ{kC%_GmWXKu^+e{$np#~^o"
    b"V*zmVuwzzYtCx3b{eG_+Id;m<_(a{<|1in8V`QBdgIBMgegE|FH_y6FvMermyCi3YGYMm&w6g2OWyTf&a8Ur>Nopg%&B7$HNtr)o"
    b"22|QCm%SJ+GkL)qU)BPVnaw?^bmM9AGUE*s-pukt$O=Z6Nuda7NqF^}P#Bu;$I;k~ciG&qQ`yXm63fi&tj<3=RrlCW#|_PO<H{md"
    b"ldJ?~oRVlov$-Y505=&oFews(B<GE@ByqFku7zPTcbSHJ&dp10sG3)<C)K7_E>lpo^VQPl1$KeaL&WU7P!H9E=q8B<fBxo&Zy)zy"
    b";jDKgdIq&hoR!Q>PnoQ<>`tiwZ9v!iz+4vJX(Ja4nP$B@xwHPI^ubn-S>WmvQVxzo?XqCBVApcLo9Na8R+dC;s@Ob?p>%CQbx{fn"
    b"%=?{e8pVP+o5}TnEVi_%SXHrzVpd}|vc&g4KYj7~PkqL@Wz9Ohg<O08b(1l|Fo17g9bN)6ne3hNPzUX2`^!U}50w3{)M1pm+$873"
    b"HUWy&IYDu>*j`oYtEUeiKUEJDaum&WxfyfzNUEJhma0k8QYwxIZ(hHAcKUNC0NDQ=iW9T5l9aspj(;*Kg0vlzd%g>m0F+s{%e|k$"
    b"zl!j@40lite`_#39?H}HP|$56-I__yV-l)(r;8*!|BZ%2-2VCL+0zC~EfQ|`Gix3H<MU$h^v&ZZr?12CHD?VpClEB}5;XD)xF&qH"
    b"d^LQHaLu@8#x*mpNqXkmLRLd_mP2#GLvs>Db3R0~Y&5e<bJj$27DaPHMRRUN8}GGgPQ7SO#b{(^aOJTiTs!94F}E9Y&6vj$bIq7*"
    b"#*uH!a&60Fv0U47ZOgSS*S0(s!!-@pG+fhgO%vv5xVGWi5!a5mcEn?exOT*&j<{wdH8E-!7+v}q!sKXpcN%V<hR>(r{AqZC8b(XQ"
    b"H`G|OG(1KPQ>P0*D;69}4DP08c>(vDWyZn+Of-uMCPc#^YVXr(cp93e1V2Tyj95faCCviD5~Nu?EFE}J+Am)kUX5nSuwdX`X%-7y"
    b"9nC^vnZW1LED`uW+B!p`SsZZ2Gz$a9P_rml5|+g*sAQ0^>{+Zq63kez1oOpYFbwpA0<Vdz{Nly?$R&g=UC7eebO}?O_tOhg`XP!D"
    b"|8plS@xEhuCnWLVNIwkmA&6T*gOBOM5Agvd?C>Fn4>#Nb7s3r6YLt^M#PH!o7g~tz_Mt@=R)lckUMS(B2r+VZg9zhS%DZ7h7csbC"
    b"!Fdt!Zme)8Y1`#%@1$z|Vk((hm!$Ox+AcSHCoS7$WWD#%C1iVf*gGlMkbT`ryt)P7<z2gUE16z4rDSrwEw!a7GswUeWzG{Nw=qfC"
    b"c`eAPV|W#|Owg$<qO9a}XkH|7%0X~uipce7^SG2eV{WpnWJjS)m=EL@(^S|$PRfYmpp0%|!EF|S42Ug=gHhP`EDCnc(wJ(BOHHMc"
    b"#Wc(+n`A}CIbHHNBT04IIAPa4iK5cO+s1hz9yU3sg``<yEJ&Phja*=W?h0|)GY+n75=R+la&44)H{Y5vFWJi_*2X0}OsOq6<4VA="
    b"C97gYi#cbLi<DE^Ho_cOQ)ys%;&aT@277j8nUt9%tZgE(U{V$=`;xfIxGtk4XS@`JmF#wrmRS<~P%ww{<0eV@6@r{Z%b9grdD*ci"
    b"L}Z2J<!Nq4U_+%0r@Ta>+xggOmZcJWB%)avIL_||6h&DIOtXX$o+fd@DS8UGV%7>8dY%_v3-XNf-WXNJD5b7N@e*PWzgQ4MjQoZJ"
    b"b;wE(Cs7&6yAeg~tv6|3nsct7c*1FLQqL%p_{bKbmn5q+7qupFu(|^qtm84UB=Nflq;jlRQ3+V{rbEOTaWW&sF}53oIab?<?${6&"
    b"Q3S!#IFiC~z69d1w?PI2Dn|?|AXJK?ID#W=<Pd3LEo(DHR2c-ClP{P$j^}BTK%{w0g2%oLz>A2b$cmT^YGOI+NJw-ruBK!!hX6qO"
    b"=L~Y5L>cc7kwsyHU6gF3=_zEGKnGyZGvXtHfS44W_o6bqV&N77k^o0D#0|J|3E2R4|I(5@*N6bYHX<rSCyB}YIzmcQ@EQ{Hixvdf"
    b"jLjXH7ZzHQrXn+27Q)m75+kTUUR3BUFH}?5FN-JXbi`|N!JY?AloxfpuOr?H84Kqi858^`5@IIqU}R-YCRNI1t){*}0K)*bYO+BI"
    b"FZnQa?8?LhBd=TJ2+$P7H;6KJEn)&7J4`V_f&ImV@WyTyX+^=FOG02~kARq1dYvcFgjiazvqFm#`zi!P_B;r;{9+6FHY6JMdo0PS"
    b"U@s^x*s;lIQr;HnjIhSeP)=wz;!9B}yB4IM?3m<)dG@u)o6ckQJjhCle@1HIZ6vvEWehRtg3Jc{v7{@hxPW9@*aI#JZ|pb{!ugdT"
    b"_=tq=Or(&9eH3_;R7X6@9BdZPGbv-|iIV<^r%LuvyklxGW%-ikru^y>36u<}UtNF$lC%>vsXJ3yvce)}ONLqG4ADpo5EH`0caAeM"
    b"{o;j_p^t-SO!h5ffs<K^t@GB&f|TOP)50O1K5?QH4>=_`x(I<*9r3Kwj4YX#ib-jG5?3&zekCR$30v{yv$V)r?%1x0B@W1=xM0ac"
    b"y%UKah(l%Rr9(@&;msAqJ_&D<lA}qhtVApYiB5E-$V@KLPC-kwL<yvbtYXlBGU+0WGBvD;bVo~mEILAZ&M=_Gsbu}7yTr0a(;Y3@"
    b"j09TQa@KrOrl>6ZZit5RvxR_?C2TY3F3(fvDU-@F7HCbiY-#8&N%<8oy32C5NOV_5uu-l}(kr%HbQh;=%QOrQTQ|BZBDQzrj7KqB"
    b"K)NHGFc!!&C?p=VNgly=x;BACQGGI(h3G!9fUPdwWf9w94FkuP873l&QZW{`U<*!nDYnjhn^_*SrKh_xW7|)m0+ws0$`ixu1KnkD"
    b"=>mx&OBf3<FubzRU2L${1YfO-d()C6$tJyoIe@Jvr!8X7yhtN{v5P<+vw2Km0eI7v1j_yi{U%SsdiS05jW@m046We?A^`!(m^nfQ"
    b"P<%CFjOEwUQV0M;C;cO#J8qJ24J$k0Ddja6`pa*@kz`24KtL8GGrSxT0)SOwt6}vacP<IO2bl%J0G26YTNZInBqtH`t`*rAUdd=h"
    b"VL+0FjoAkyh0IcjiC|1-z#)g$R9gbfq)}p7lem<HmAqHW{wS$KnnMW)GFZ5ALvcVM3u794KPG`yf)XWBPI%@WLjn<?MLibmf79wo"
    b"`i^CjbPx7`ZsMH%EAl*O1@Y?}8F%*S<E+eM*b=f%5wBt7lEkcGG2{Zt#$X{N8)LQ!FUmA9?-Rvn1JWhyWf`+rJ=y2RgIKok3wY7E"
    b"A+_NEkX%JJ#Q;Rub0B8X!pQL=)_qPIMDP_k79eI?S^3$xp%?Zc3F4%|?Anw5djE=)Br5rRNm{NsfkQT1Oa_Gr>`&umMOs<dFQ?Uq"
    b"U&*B9i-S50-FoXsCL&`cB)_+mxB@;m`}t%FBlfwYlyr(cY-4$CA@h}rZ%E5RQ3O9ZC)>qtGzP>zFv(xuN{tO$YckA<_{kX|k9}4$"
    b">6U$x6j%!T!=?BbSSZ;aBEBaaGl0>S{dZdB6ZWN{T<qJ@(#dN%ErL<-0BOaDId-D;GLsE-GJzQ+OQ<XoPON3TRFde}`K5;mqzXgm"
    b"7c^QlikKZp8$n?+Xf<zr7m{>fwf73@Rmp?C4^IyyFLE!OiDRLktQ(hWDkNaF{Dku$w$j(uo?QF(wC_N>9-Y(WhN0cqT$Ah{xzTGs"
    b"9qs3*{e-pSfZKL*Y|)Na@O^}N+H+HT_G-^>9VDQ=P{MYa$Vz(=@>ZNkRR=-qAbTC)paVj*2aR_`O7E{ah>0U%-PWZ8q7K;90j@eA"
    b"SbI>rHa*t0R{`8Pt<XVHbkH3g)JX>o(_ZP^BIk8cdlmJ9jAStuMb<&Lbx?O5EP(b#!7rT7R%mZRc;OVQqJyE)!S?82j&!h4+8Zs`"
    b"rnhC<n>gG#tEYn@)WJ6DU@mpApxPT%he&U0wKu`Ib5>agL#>1D*1??XVBxhl`hAp3NFhrE_9ES?p~FI=!`h<5a-+ivr2V4gHaY!T"
    b"rTr4-m58bQLZ`#pr^9ln!wRYWqUqZ7YpM22EbiQ@tHVO9!`iIFa;?J(uKl9!+VpF^_AUVKoL8WOqo9NDpo2T1gNLEL)8WD^?~7>f"
    b"q9D?8UW*Qnj1Iny4(^T)9+37<k!#cYO4_?nh}E1|rGulTgYTt-JEnt&roGeV9lLH1PFvYgmMvx8-?sv80lR?TdDRSc!*z`8IhTC("
    b"JM=bkK09B<lNS>p{ylF@^EN(zF#*#cF@nSl5<>(7A~8iQ(Z;09ynHC<vM<}RC(pv0_dd_k&#>4o!xC~UY${z|#T(0hR;BnrKE+J-"
    b"XHgQDKk-pz)XSmdz3d6M;xZ>uFK3eVawf6MmbmOlDY+85WFb#tLyDvfDUvi~N9;~^q}Y=fF(1l_bg7U4)O<G+VnQawB|)IaApz3m"
    b"K0J2rrar{f?q)n<pX?BhiKv5In6(p^=O|r@BXubb%jpf<fA1xiSF%hmVW0J0zcTK%E6HVuGg!EHSNge(eA4AG{3}(GzmNdMzeW|i"
    b"*QfICg#gZ5?B*<FWM;CPuP9uu!oLcYx;%xv{v=z0*5xStD^F!ePsF?LHx=QnChwQZ3r%Twjj4Ns$%glq?xZ1%WFUmTmxU-MJIROc"
    b"!>}&z(0vK!_e13yF!%Y_PPSoRqQTz(t}72;bD7<5xHA9cmj7tWf2)<s7h3N1nIe3kWy5D#?u#t{A(q+w{3?@=u3U@|^**=audF1|"
    b"tbe<|yDteGs`>2ZDEb3`T#XNK&km`IQure9%`ugo{6VeG&+%EW{~YjZu5fmA6dkHXU7h3eLn*{=Ci3C@d9H8Z0g;~&fv29xw~17u"
    b"9vq=<-<-MIT=Jpj9+;C4H}}B&#vB_Ari*H`QBUUWYyNOvzRw&Ehu?nxM4!(ae6dWNYq(IY-r+U7%i4d~9zOCOEm!<KCrxV8k__g{"
    b"i}3+}Ad2QrPdAHIyE#3jiH?S|;Y1DBtIcpMrTlg1>3lhFPftgidT~BcvvK!*Hl7^Jh;p|3h+3eBQP&2>=;)U{`gdH*|3X93Vt6DC"
    b"`8LDj#gd*b>g8xQ4$p9|!zn%`Z}8pesNW6#_$&e%=rNjn^hm`g5ctUd8mlkSD$6w5A;cI*@6?xHDl=C9sYWAiiD*W-($c{1RX!f$"
    b"vsE+RAw?O1!|Y>De>(2-+IQ}ATqwxtb+!0q^k$;Ynu)rqZuo<DnGqo8*@j8eu*0lPsH`5)XgU@EC`Z)T0ZXDc&Q9UNqkDSi3y5Ld"
    b"4)wmv+#O9ky!Y+Bn;&;W<6_0U*SMdWQ_e6xlE?=-Oa9Kzz=~h(5S8Rx=L&}e{0QBD)-<bTG`#E?E>^Si#p)UwA|Fr7H~yn{6aN0+"
    b"@%*zpCn{W37o<enS*;)8aD{r-9j`EOlL9rr7PVZpKpY^(&ooJr(Y32MRONwS^1ntS1Aub>T)qF*Hh<my&JR~kB{}6&5=H<QULczA"
    b"nGF;5hD;t^`GkgD(2>26<3-eJS*z}BjQO%%aTCW`a-v54V;sf@>NSsGqn7oxKNx~759<P^cGL$ucizXln2$Nb(X-h@nCI$P9=O+2"
    b"U-ShsOrGx_a7)rA4>&r38d8zeS6)%CFYyiknhZjg^i19ESeZo_Yc}qy#|FxgAF927pY0q`v1^@ZPTdn6?$9@kXm}UsQX}wUd{;pj"
    b"La4+b+J~Uf(ml9OS532Bx4k1UhF$w>s0)?!$F{~90DWHVZx`*YD_XSrT<A6Z_W~jaaDDO*;37=I^mT+`@Y}mQ{4Te4M{@L!{OL}|"
    b"=pWP(&f_4`5vq_0AxnN^)9erjFJ#rx=55m=7Pv99^2i(%NuJ|fVuT6NjV)(K+8$uf$HEh(S|+7hQDNf}38e9F?V_*&+%#1Yj*sXY"
    b"otrW6S9n7irdN!-69E-H_I0tC-C!-7$y)TxdIg>x+Y>18adsjYXiX@S?)g^gX{o2aV3s`|!y-Pj<DliPMI$wH*WzZU<Uz0GZC8>C"
    b"P3%X}mBd1mx{_@$J@ojkJbUM#K_6Rr_EY!lfq(Wa>Y9zD*~B&b%<L5DUeWh~is&DN34d>Tk2^m6fTtbfw!*Bh%CQvxdAIm26*IiP"
    b"a^<gmqd!nre|FEpTwYU0Kgq*4-KY#^Xk7nqOz*kwp8vSp@w-llT>NUccrC>yb;Un+WB8F;eeDME>W6>xk6+PajQK}j=HMzA{2rY9"
    b"iL3|zdV^oDyPw9<x!>Po;MDYZvSasr*#jvHw7O6GCX(U#rXuFpiBpvMY}&d78g}k<S+(SLzn$Pnn2EZfVC|OvThH6+rBkD8pcjbv"
    b"g6wc`45yu6SoB=M>1q+Q&bNO&r2!X2-k^=y+um={%ckGf$>7dXQ!Q_)sl~B5Jh|g`-=n?10$p$Vy&qwGG?kAyfK35$_N)m}jQ`nQ"
    b"o8q`R8T<-yv%03LMA{HpaX4zk1Ec$Hyl41ggS@j_z^6eWONxHnMEm$gZSLV4iTW5V9ddItPK*(e$7XcgJ23~uB=+57bJ+G3A7S9?"
    b"4*wb)8FZI_(~tAd50J(-J=s0RXLUTiJ0yqfv+aB_J9S52jo$iz=A3Nomwo5tZE(iqE2tx+z9aSUA(gn=sXJ@PB34(HX8cxcvb+ON"
    b"0Pm#9<D}V)_<NX%a^H9KX4Q5Co@n09KTGZ~;4L~B$=JT&=8xyF)M+%vAWanAuWoqc8%)%VJWbuBM|{r@-)?p&+k<MohI=FH8`b?n"
    b"`Om$iOOCmCoh2LENzom$<*3^oV<%<1@c>Ss8qKP<y65jx;KPSzkwH4cuE;2yPB&KTMYY7YNC0NlG}SE*wfkwJUO4j<oP7#^=F^F~"
    b"HWPP98nzqcII4r~a$S+B8$L=lFJHbKlBuI%93L9L))SAIT>XwTKu!$oY2-aOQ}jy4jMKP=-6}r1Hei%Z1eh)s<XFlVSqbAqG<exB"
    b"G<OPLbZr_VZRV~~EJRe&f8^bA3e}FN{L3RL<B*S^z1Z{YEiQzo@yDioR?TF-a=_v+@a+;jGdM-_9pBw;<hxb-631{|)yuY?xz;w@"
    b"GdY~gH9ESS2=$o9zuB(g$=0*38}_ju$%ae^q5WeenJ=e&5Fj`1HmO`&SJUk&8FSXIpNP#lt5Xrm$WU?Fi#uhrx4gS?$GQw8vZ9i("
    b"g~sDMHrXMD?C`a|CbUA%q4zKF(Lm_W+foO1G86&sqxCiH`{nrV?o;g6NW>MgHp@rWolUyB?P~k3dW_>LH@j6Zwmr)$PGRi#eBV|h"
    b"kis^Pd{jjOV-jrRh|!TxobHh!>Vn}idq~QAwco$L>x1lUg@F;rXatMJy-kSs?<f5#VJCt|Eb-s?UYq2%e5}h~Gvd8wj{oxVIhj)o"
    b"?frYYxmT~}k~wdzJ&yLU#hgqA>cf8mAwNfwuORJK{#{@1m}#jT>?xzv8Eq-TQG$H%?dWb*H&Z^P_0`L7lsob7B`R~t&ubWQTpEN#"
    b"LNDrkbwLlf=E1VQ9t~eVe?m5Sgr?p4vc~=+)jNRZj>ZAfJw(6g+W>$FQ(Rx~z`KX;pYDhTU;h2PJ|;@?ns%izMQb$l{e@6=n$q4v"
    b"l>JPP<%_uEX|r2=;fhBi_npV_^z6h}f6nz~nay^Qc-C*0iFWsZZ;o%sVF$EEjF^TzfM~}9X!;I7^I-rn`Nci*i$3|quaKYJBR}ht"
    b"pY0%!Wb6RvPPx3}A^4-6*)R1cZ{iSGzy6Mz>RP$eEgL>>vpb)2j8kMQ2T!2Q`wFuC%%c@gAUT1CxusvmiGl;1uQ_s$!&(IDFZ=t`"
    b"`#c00ogDkA`0Qubz<r%N?aWp1k&da8!;?MA`VozEjJ%?`kE-V4-LI-BRpOPSk`t?;ETBWI5W<IpO<OhX28XVWhFTBZs>8>M9;u-m"
    b"Fgo<4o<&1<m}w~Jn8nANPW<_%M?8owFR(eds(z`dWHXW)6II{LH|^=_m!mh$w(b;TcmI+js@00@=IO*8qD?(hbG*&ReePM;&N;8K"
    b"VP9j!RC^5gx@xg>Q(L$@_}hxuD_xy0@2?nbbGv^H;A3l7<SUQR%tQ^Fv*8#A`<@?uP@lw7dd)bO?Z@%`PV1X#y>8Vr{s&+1Q?vE{"
    b"X4lP}JeT3-m-$+KzB#1-hU#<W##HZVTjKd;J6~?=J(-Y2k2UlQ|GLJ6>HDI(yWI4W_76Af#k{R=+7UH4<_)Tkt$6-#G{o7qL%@&G"
    b"$yz+bX}-ff)7n9r>z95ktE(F08GjIWJ=q*=s&`y_qTKPyM`b$R9YDX@c<=#ftbU>U&o_vD=sNm*cJujc+>;A;!7Y(bNet=zFeiH^"
    b"x7^JMx?Vn?i5VPi#@(sP=%oi^u{B<e;ci1e$Gfq!w0C+sTTM?-_tJJ?2RV;;AKpKPZ@~$Zo0HwaJo54FJ;1o$1T^a32Y(MX2LF^a"
    b"J#Y63FU<Yd<Tn`1kpu_$kR>O(dU^^oetPQSP`2rtTZ*<}*9ZR%k7y&{"
)


def _load_ajpg_decoder():
    try:
        import ajpg_to_png as mod
        return mod, "file", ""
    except Exception as e:
        file_err = f"{type(e).__name__}: {e}"
    try:
        src = zlib.decompress(base64.b85decode(_AJPG_EMBEDDED)).decode("utf-8")
        mod = types.ModuleType("ajpg_to_png")
        mod.__dict__["__file__"] = "<embedded in playtime_log_gui.py>"
        exec(compile(src, "ajpg_to_png.py", "exec"), mod.__dict__)
        return mod, "embedded", ""
    except Exception as e2:
        return None, "unavailable", (f"import of ajpg_to_png.py failed "
                                     f"({file_err}); embedded fallback also "
                                     f"failed ({type(e2).__name__}: {e2})")


ajpg_to_png, AJPG_SOURCE, AJPG_IMPORT_ERROR = _load_ajpg_decoder()
AJPG_AVAILABLE = ajpg_to_png is not None

# --- optional drag & drop ----------------------------------------------------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

APP_TITLE = "Wii Message Board Converter"
APP_VERSION = "a1.3"
TS_TOKEN = "(yyyy-mm-dd-hh-mm-ss)"
DEFAULT_OUT_NAME = ("Wii Message Board - Playtime + Message Export "
                    + TS_TOKEN + ".xlsx")
FILETYPES = [("Wii dumps", "*.zip *.rar *.vff *.VFF"), ("All files", "*.*")]

ALT_TEXT = ("A vff (raw) dump is preferred: it contains everything a cdb "
            "dump/cdbackup.vff have, plus title IDs and photos.\n\n"
            "Two dumps from different consoles will merge, with any "
            "duplicates counted only once.")


def program_dir():
    """Directory the running program lives in (the script's folder, or the
    folder containing the .exe for a PyInstaller build)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


DEFAULT_WORKBOOK_DIR = os.path.join(program_dir(),
                                    "Extracted Playtime Spreadsheets")
DEFAULT_PHOTOS_DIR = os.path.join(program_dir(), "Extracted aJPG Photos")

# --- persistent cache: parsed dumps + decoded photos -------------------------
# Lives in a "cache" folder next to the program (the .exe's folder for a
# PyInstaller build), so re-scanning an unchanged dump skips the parsing and
# AJPG-decoding work entirely.  Scan entries are invalidated automatically
# whenever the app version or the converter script changes.
CACHE_DIR = os.path.join(program_dir(), "cache")

_CACHE_MISS = object()


def _dump_cache_key(path):
    """Stable cache key for a dump input: sha1 of the file contents, or of
    the (relative name, size, mtime) listing for a folder.  None if the path
    cannot be read (caching is then simply skipped)."""
    h = hashlib.sha1()
    try:
        if os.path.isdir(path):
            for root, _dirs, files in os.walk(path):
                for fn in sorted(files):
                    fp = os.path.join(root, fn)
                    st = os.stat(fp)
                    h.update(os.path.relpath(fp, path)
                             .encode("utf-8", "replace"))
                    h.update(str(st.st_size).encode())
                    h.update(str(st.st_mtime_ns).encode())
        else:
            with open(path, "rb") as f:
                for chunk in iter(lambda: f.read(1 << 20), b""):
                    h.update(chunk)
    except OSError:
        return None
    return h.hexdigest()


def _converter_stamp():
    """Baked into every cache file name so cached scans from an older app
    version or converter script are never reused."""
    try:
        mod = sys.modules.get("playtime_log_to_excel")
        mtime = os.path.getmtime(mod.__file__) if mod else 0
    except Exception:
        mtime = 0
    return f"{APP_VERSION}-{int(mtime)}"


def _cache_path(kind, key):
    return os.path.join(CACHE_DIR,
                        f"{kind}-{key}-{_converter_stamp()}.pkl")


def _cache_read(kind, key):
    try:
        with open(_cache_path(kind, key), "rb") as f:
            return pickle.load(f)
    except Exception:
        return _CACHE_MISS


def _cache_write(kind, key, obj):
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        path = _cache_path(kind, key)
        tmp = path + ".tmp"
        with open(tmp, "wb") as f:
            pickle.dump(obj, f, pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, path)   # atomic: never leave a half-written entry
    except Exception:
        pass


def load_source_cached(path):
    """load_source() with the per-dump disk cache - used by the convert and
    export workers when the current scan results can't be reused directly."""
    key = _dump_cache_key(path)
    if key:
        b = _cache_read("scan", key)
        if b is not _CACHE_MISS:
            return b
    b = load_source(path)
    if key:
        _cache_write("scan", key, b)
    return b


def decode_photo_image(it):
    """Decode one attached photo to a PIL image, using a PNG cache in
    cache/photos keyed by the photo bytes (sha1).  Returns (image, from_cache)."""
    cpath = os.path.join(CACHE_DIR, "photos",
                         hashlib.sha1(it.ajpg).hexdigest() + ".png")
    if os.path.exists(cpath):
        try:
            with Image.open(cpath) as im:
                return im.copy(), True
        except Exception:
            pass
    img = ajpg_to_png.convert(it.ajpg)
    for _attempt in range(2):   # one retry guards against flaky filesystems
        try:
            os.makedirs(os.path.dirname(cpath), exist_ok=True)
            tmp = cpath + ".tmp"
            with open(tmp, "wb") as f:
                img.save(f, "PNG")
            os.replace(tmp, cpath)   # atomic: never a half-written PNG
            break
        except Exception:
            pass
    return img, False


class ToolTip:
    """Hover tooltip for a widget (plain tkinter, no dependencies)."""

    def __init__(self, widget, text):
        self.widget, self.text, self.tw = widget, text, None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, _e=None):
        if self.tw is not None:
            return
        x = self.widget.winfo_rootx() + 18
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tw, text=self.text, background="#ffffe0",
                 relief="solid", borderwidth=1, padx=6, pady=4,
                 wraplength=320, justify="left").pack()

    def hide(self, _e=None):
        if self.tw is not None:
            self.tw.destroy()
            self.tw = None

CHECK_ITEMS = [
    ("logs",     "Play-time logs → Combined Totals + per-day sheets"),
    ("achv",     "Achievements (medals, records, pro status, stamps, trophies)"),
    ("letters",  "Letters (Animal Crossing, Nintendo mail)"),
    ("personal", "Personal notes (hand-written board messages)"),
    ("titleids", "Game title + Title-ID matching (vff dumps)"),
    ("photos",   "Attached photos (AJPG - proprietary Wii format)"),
    ("arcs",     "Letter stationery (U8 archives)"),
]

# short label + unit per checklist key, used in the default (non-verbose) view
CHECK_SIMPLE = {
    "logs":     ("Play-time logs", "days"),
    "achv":     ("Achievements", "memos"),
    "letters":  ("Letters", "letters"),
    "personal": ("Personal notes", "notes"),
    "titleids": ("Game title + Title-ID matching", "title IDs"),
    "photos":   ("Attached photos", "photos"),
    "arcs":     ("Letter stationery", "archives"),
}


def convert(bundles, out_dir, out_name, log=print):
    """UI-independent conversion. bundles: list of SourceBundle.
    Returns the output path on success, raises on failure."""
    if not out_name.lower().endswith(".xlsx"):
        out_name += ".xlsx"
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, out_name)
    with contextlib.redirect_stdout(io.StringIO()) as buf:
        build_workbook(bundles, out_path)
    for line in buf.getvalue().splitlines():
        log(line)
    return out_path


def fmt_hhmm(minutes):
    return f"{minutes // 60}:{minutes % 60:02d}"


# ===========================================================================
# AJPG photo collection & decoding
# ===========================================================================

class PhotoItem:
    """One attached AJPG photo plus the message it accompanied."""
    __slots__ = ("date", "title_id", "msg_title", "msg_body", "source",
                 "dump", "ajpg", "image", "key")

    def __init__(self, date, title_id, msg_title, msg_body, source, dump,
                 ajpg, key):
        self.date = date
        self.title_id = title_id
        self.msg_title = msg_title
        self.msg_body = msg_body
        self.source = source
        self.dump = dump
        self.ajpg = ajpg
        self.image = None       # PIL.Image once decoded
        self.key = key          # memo source path this photo belongs to

    @property
    def date_str(self):
        return self.date.strftime("%Y-%m-%d %H:%M") if self.date else "?"


def find_ajpgs(data):
    """Yield (offset, size) of every AJPG image embedded in a record blob."""
    off = 0
    while True:
        off = data.find(b"AJPG", off)
        if off < 0:
            return
        if off + 16 <= len(data):
            w1, size = int.from_bytes(data[off + 4:off + 8], "big"), \
                int.from_bytes(data[off + 8:off + 12], "big")
            q, w, h = (w1 >> 24) & 0xFF, w1 & 0x7FF, (w1 >> 11) & 0x7FF
            if 1 <= q <= 100 and 1 <= w <= 0x7FF and 1 <= h <= 0x7FF \
                    and 16 <= size <= len(data) - off:
                yield off, size
        off += 4


def collect_photos(bundle):
    """Extract every attached AJPG photo from a scanned dump, with the text of
    the message it was attached to. Returns a list of PhotoItem (undecoded)."""
    items = []
    try:
        _label, acc, kind = open_source(bundle.path)
    except Exception:
        return items
    names = acc.namelist()
    if kind == "vff-records":
        # raw .000 board records; a photo lives inside the message's own record
        for name in names:
            norm = name.replace("\\", "/")
            m = VFF_PATH_RE.search(norm)
            if not m:
                continue
            try:
                data = acc.read(name)
            except Exception:
                continue
            if b"AJPG" not in data:
                continue
            rec = parse_cdbfile(data)
            for off, size in find_ajpgs(data):
                items.append(PhotoItem(
                    date=vff_record_date(rec, m) if rec else None,
                    title_id=m.group(6).split("_")[0].upper(),
                    msg_title=rec["title"] if rec else "",
                    msg_body=" ".join(rec["body"].split()) if rec else "",
                    source=norm, dump=bundle.label,
                    ajpg=data[off:off + size], key=norm))
    else:
        # cdb text export: '<msg> ripl_board_record-attachment.ajpg' sidecars
        by_lower = {n.lower(): n for n in names}
        for name in names:
            low = name.lower()
            if not low.endswith("-attachment.ajpg"):
                continue
            try:
                data = acc.read(name)
            except Exception:
                continue
            if not data.startswith(b"AJPG"):
                continue
            txt = by_lower.get(low[:-len("-attachment.ajpg")] + ".txt")
            date, title, flat = None, "", ""
            if txt:
                try:
                    memo = parse_memo_text(acc.read(txt).decode("utf-8", "replace"))
                    date = parse_board_date(memo["date"])
                    title, flat = memo["title"], memo["flat"]
                except Exception:
                    pass
            items.append(PhotoItem(
                date=date, title_id="", msg_title=title, msg_body=flat,
                source=name.replace("\\", "/"), dump=bundle.label,
                ajpg=data, key=(txt or name).replace("\\", "/")))
    return items


def dedup_photos(items):
    """Drop duplicate copies of the same attached photo (same image appearing
    at the same message date - i.e. two dumps of the same console)."""
    seen, out = set(), []
    for it in items:
        h = (hashlib.sha1(it.ajpg).hexdigest(), it.date_str)
        if h not in seen:
            seen.add(h)
            out.append(it)
    return out


# ===========================================================================
# Dependency self-check + click-to-install
# ===========================================================================
# (import name, pip package, required, purpose)
DEPENDENCIES = [
    ("openpyxl",    "openpyxl",    True,  "spreadsheet (.xlsx) output"),
    ("ttkbootstrap", "ttkbootstrap", False, "themed UI + dark mode"),
    ("PIL",         "pillow",      True,  "attached-photo decoding & viewing"),
    ("tkinterdnd2", "tkinterdnd2", False, "drag & drop of dump files"),
    ("rarfile",     "rarfile",     False, ".rar dump input (also needs "
                                          "WinRAR/UnRAR installed)"),
]


def _augment_sys_path():
    """pip may install into the user site-packages directory; if that
    directory did not exist when this process started it is missing from
    sys.path, so add it (and the global site dirs) on demand."""
    import site
    cands = [_EXTRA_PKG_DIR, site.getusersitepackages()]
    with contextlib.suppress(Exception):
        cands.extend(site.getsitepackages())
    for p in cands:
        if isinstance(p, str) and os.path.isdir(p) and p not in sys.path:
            sys.path.append(p)


def dep_installed(import_name):
    try:
        _augment_sys_path()
        importlib.invalidate_caches()  # pick up packages installed this session
        return importlib.util.find_spec(import_name) is not None
    except Exception:
        return False


class DepsDialog(tk.Toplevel):
    """Shows the status of every dependency; missing ones get an Install
    button that runs 'sys.executable -m pip install <pkg>'."""

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        self.title("Dependencies - " + APP_TITLE)
        self.transient(app.root)
        self.resizable(False, False)
        if TTB_AVAILABLE and app.dark_var.get():
            self.configure(background="#222222")
        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Packages this app uses:",
                  font=("", 10, "bold")).pack(anchor="w", pady=(0, 6))
        self.grid_frame = ttk.Frame(body)
        self.grid_frame.pack()
        ttk.Button(body, text="Re-check", command=self.refresh).pack(pady=(10, 0))
        self.refresh()

    def _row(self, r, name, purpose, status, color, install=None):
        ttk.Label(self.grid_frame, text=name, font=("", 10, "bold")
                  ).grid(row=r, column=0, sticky="w", padx=(0, 10), pady=2)
        ttk.Label(self.grid_frame, text=purpose, wraplength=240, justify="left"
                  ).grid(row=r, column=1, sticky="w", padx=(0, 10))
        ttk.Label(self.grid_frame, text=status, foreground=color,
                  wraplength=200, justify="left"
                  ).grid(row=r, column=2, sticky="w", padx=(0, 10))
        if install:
            ttk.Button(self.grid_frame, text="Install", command=install
                       ).grid(row=r, column=3, sticky="w")

    def refresh(self):
        for w in self.grid_frame.winfo_children():
            w.destroy()
        frozen = getattr(sys, "frozen", False)
        for r, (mod, pipname, required, purpose) in enumerate(DEPENDENCIES):
            ok = dep_installed(mod)
            if ok:
                status, color = "✓ installed", "#087008"
            elif required:
                status, color = "✗ MISSING (required)", "#b00020"
            else:
                status, color = "✗ missing (optional)", "#a66a00"
            install = (None if ok else
                       (lambda p=pipname: self.app.install_dep(p, self.refresh)))
            self._row(r, pipname, purpose, status, color, install)
        n = len(DEPENDENCIES)
        ttk.Separator(self.grid_frame).grid(row=n, column=0, columnspan=4,
                                            sticky="ew", pady=6)
        conv = ("✓ loaded" if CONVERTER_AVAILABLE
                else "✗ " + (CONVERTER_IMPORT_ERROR or "not found"))
        self._row(n + 1, "playtime_log_to_excel.py",
                  "converter core - keep next to the GUI script",
                  conv, "#087008" if CONVERTER_AVAILABLE else "#b00020")
        if AJPG_AVAILABLE:
            aj = ("✓ external ajpg_to_png.py" if AJPG_SOURCE == "file"
                  else "✓ built-in embedded copy")
            self._row(n + 2, "AJPG photo decoder",
                      "decodes the Wii's proprietary 'AJPG' photo attachments "
                      "(an embedded copy is built in as fallback)", aj, "#087008")
        else:
            self._row(n + 2, "AJPG photo decoder",
                      "decodes the Wii's proprietary 'AJPG' photo attachments",
                      "✗ " + AJPG_IMPORT_ERROR, "#b00020")
        if frozen:
            ttk.Label(self.grid_frame, wraplength=560, justify="left",
                      foreground="#a66a00",
                      text="This is a packaged .exe - 1-click installs use a "
                           "system Python (the 'py' launcher or 'python' in "
                           "PATH) and land in a 'python_packages' folder next "
                           "to the .exe. Alternatively rebuild with the "
                           "updated Build Executable.cmd so it gets bundled."
                      ).grid(row=n + 3, column=0, columnspan=4, sticky="w",
                             pady=(8, 0))


class InputRow:
    """One dump selector: path entry + file/folder/clear buttons + drop."""

    def __init__(self, app, parent, row, label, tooltip=None):
        self.app = app
        self.var = tk.StringVar()
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w")
        self.entry = ttk.Entry(parent, textvariable=self.var, width=34)
        self.entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=1)
        if tooltip:
            info = ttk.Label(parent, text="ⓘ", foreground="#005a9e",
                             cursor="question_arrow")
            info.grid(row=row, column=3, sticky="w", padx=(3, 0))
            ToolTip(info, tooltip)
        btns = ttk.Frame(parent)
        btns.grid(row=row + 1, column=0, columnspan=3, sticky="w", pady=(2, 4))
        ttk.Button(btns, text="File…", width=8, style=_bs("primary"),
                   command=self.browse_file).pack(side="left", padx=(0, 5))
        ttk.Button(btns, text="Folder…", width=8, style=_bs("primary"),
                   command=self.browse_folder).pack(side="left", padx=(0, 5))
        ttk.Button(btns, text="Clear", width=8, style=_bs("primary"),
                   command=self.clear).pack(side="left")
        if DND_AVAILABLE:
            self.entry.drop_target_register(DND_FILES)
            self.entry.dnd_bind("<<Drop>>", self.on_drop)
        # NOTE: no auto-scan on edit - the user presses "Scan Dumps" when ready

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select a Wii dump (.zip, .rar, .vff)", filetypes=FILETYPES)
        if path:
            self.var.set(path)

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select a cdb or vff folder")
        if path:
            self.var.set(path)

    def on_drop(self, event):
        paths = self.entry.tk.splitlist(event.data)
        if paths:
            self.var.set(paths[0])

    def clear(self):
        self.var.set("")
        self.app.start_scan()   # empty inputs -> reset the view

    def get(self):
        return self.var.get().strip().strip('"')


class SortableTree(ttk.Treeview):
    """Treeview whose columns sort (asc/desc) on header click."""

    def __init__(self, parent, columns, widths):
        super().__init__(parent, columns=columns, show="headings")
        self._sort_keys = {}
        self._sort_rev = {}
        for col, w in zip(columns, widths):
            self.heading(col, text=col, command=lambda c=col: self.sort_by(c))
            self.column(col, width=w, anchor="w")
        sb = ttk.Scrollbar(parent, orient="vertical", command=self.yview)
        self.configure(yscrollcommand=sb.set)
        self._scrollbar = sb

    def pack_with_scroll(self, **kw):
        self._scrollbar.pack(side="right", fill="y")
        self.pack(**kw)

    def set_sort_key(self, column, func):
        """func(row_values_tuple) -> sortable value for the given column."""
        self._sort_keys[column] = func

    def sort_by(self, column):
        rev = self._sort_rev.get(column, False)
        keyfunc = self._sort_keys.get(column)
        rows = [(self.set(item, column), item) for item in self.get_children("")]
        if keyfunc:
            rows.sort(key=lambda kv: keyfunc(kv[0]), reverse=rev)
        else:
            rows.sort(key=lambda kv: kv[0], reverse=rev)
        for idx, (_v, item) in enumerate(rows):
            self.move(item, "", idx)
        self._sort_rev[column] = not rev

    def refill(self, rows):
        self.delete(*self.get_children(""))
        for row in rows:
            self.insert("", "end", values=row)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.queue = queue.Queue()
        self.bundles = []
        self.bundles_key = None
        self._scan_job = None
        self._scan_gen = 0
        self._busy = False
        self._eff_logs = []
        self._eff_memos = []
        self.photo_items = []          # decoded PhotoItem list
        self._memo_photo_count = {}    # memo source path -> attached photo count
        self._photo_by_key = {}        # memo source path -> PhotoItem
        self._thumb_refs = []          # keep ImageTk.PhotoImage alive
        self._detail_img_ref = None
        self._gallery_btns = []
        self._sel_photo = None
        self.show_log_var = tk.BooleanVar(value=False)   # Toggle menu: Log
        self.verbose_var = tk.BooleanVar(value=False)    # Toggle menu: Checklist
        self.dark_var = tk.BooleanVar(value=False)       # Toggle menu: Theme
        self._check_rows = None        # filled by _set_checklist

        root.title(f"{APP_TITLE}  ({APP_VERSION})")
        root.minsize(900, 600)

        # optional themed UI (ttkbootstrap): match the OS light/dark setting
        self._ttb_style = None
        if TTB_AVAILABLE:
            self.dark_var.set(system_prefers_dark())
            self._ttb_style = ttkbootstrap.Style(
                theme="darkly" if self.dark_var.get() else "flatly")
            self._fix_tree_colors()

        # ---- top toolbar ---------------------------------------------------
        bar = ttk.Frame(root)
        bar.pack(fill="x")
        self.toggle_btn = ttk.Button(bar, text="◀", width=3, command=self.toggle_sidebar)
        self.toggle_btn.pack(side="left", padx=(6, 4), pady=4)
        ttk.Label(bar, text=APP_TITLE, font=("", 11, "bold")).pack(side="left")
        self.status = tk.StringVar(value="Ready")
        ttk.Label(bar, textvariable=self.status, foreground="#555").pack(
            side="right", padx=8)

        # ---- bottom progress bar (text drawn inside the bar) ---------------
        pf = ttk.Frame(root)
        pf.pack(fill="x", side="bottom")
        ttk.Style(root).configure("Tall.Horizontal.TProgressbar", thickness=18)
        self.progress = ttk.Progressbar(pf, mode="determinate", maximum=100,
                                        style="Tall.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=6, pady=(0, 4))
        self.progress_text = ttk.Label(self.progress, text="", anchor="center",
                                       font=("", 8))
        self.progress_text.place(relx=0.5, rely=0.5, anchor="center")

        # ---- body: sidebar + main -----------------------------------------
        body = ttk.Frame(root)
        body.pack(fill="both", expand=True)

        self.sidebar = ttk.Frame(body, width=340)
        self.sidebar.pack(side="left", fill="y", padx=(6, 0))
        self.sidebar.pack_propagate(False)
        self._build_sidebar()

        main = ttk.Frame(body)
        main.pack(side="left", fill="both", expand=True, padx=6, pady=(0, 6))
        self._build_main(main)

        self._build_menubar()
        self._restyle_plain_widgets()   # match Text panes to a dark OS theme
        self._apply_log_visibility()
        self._set_checklist(None)   # render checklist in the current mode

        if not DND_AVAILABLE:
            self.log("tip: install tkinterdnd2 for drag && drop support")
        self.root.after(100, self._poll_queue)
        self.root.after(400, self._startup_dep_check)

    # ------------------------------------------------------------- sidebar
    def _build_sidebar(self):
        sb = self.sidebar

        inp = ttk.LabelFrame(sb, text="Input dumps (folder / .zip / .rar / .vff)")
        inp.pack(fill="x", padx=4, pady=4)
        inp.columnconfigure(1, weight=1)
        self.input1 = InputRow(self, inp, 0, "Dump 1 (vff recommended):",
                               tooltip=ALT_TEXT)
        ttk.Separator(inp).grid(row=2, column=0, columnspan=3, sticky="ew", pady=2)
        self.input2 = InputRow(self, inp, 3, "Dump 2 (optional):")

        ttk.Button(sb, text="Scan Dumps", style=_bs("primary"),
                   command=self.start_scan).pack(fill="x", padx=4, pady=(0, 4))

        chk = ttk.LabelFrame(sb, text="Expected workbook contents")
        chk.pack(fill="x", padx=4, pady=4)
        self.check_vars = {}
        self.check_labels = {}
        for key, text in CHECK_ITEMS:
            var = tk.StringVar(value="—  " + text)
            self.check_vars[key] = var
            lbl = ttk.Label(chk, textvariable=var, foreground="#777",
                            wraplength=300, justify="left")
            lbl.pack(anchor="w", padx=4, pady=1)
            self.check_labels[key] = lbl
        self.scan_status = tk.StringVar(value="Select dump(s), then press Scan Dumps.")
        ttk.Label(chk, textvariable=self.scan_status, foreground="#005a9e",
                  wraplength=300, justify="left").pack(anchor="w", padx=4, pady=(4, 2))

        out = ttk.LabelFrame(sb, text="Output")
        out.pack(fill="x", padx=4, pady=4)
        out.columnconfigure(0, weight=1)
        ttk.Label(out, text="Workbook folder:").grid(
            row=0, column=0, columnspan=2, sticky="w")
        self.out_dir = tk.StringVar(value=DEFAULT_WORKBOOK_DIR)
        ttk.Entry(out, textvariable=self.out_dir, width=30).grid(
            row=1, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(out, text="Folder…", width=7, style=_bs("primary"),
                   command=self.browse_out).grid(row=1, column=1)
        self.out_name = tk.StringVar(value=DEFAULT_OUT_NAME)
        ttk.Entry(out, textvariable=self.out_name, width=30).grid(
            row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        self.convert_btn = ttk.Button(out, text="Convert to spreadsheet",
                                      style=_bs("primary"),
                                      command=self.on_convert)
        self.convert_btn.grid(row=3, column=0, columnspan=2,
                              sticky="ew", pady=(4, 2))
        ttk.Separator(out).grid(row=4, column=0, columnspan=2,
                                sticky="ew", pady=4)
        ttk.Label(out, text="Photo folder:").grid(
            row=5, column=0, columnspan=2, sticky="w")
        self.photo_out_dir = tk.StringVar(value=DEFAULT_PHOTOS_DIR)
        ttk.Entry(out, textvariable=self.photo_out_dir, width=30).grid(
            row=6, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(out, text="Folder…", width=7, style=_bs("primary"),
                   command=self.browse_photo_out).grid(row=6, column=1)
        ttk.Button(out, text="Export pictures to .png", style=_bs("primary"),
                   command=self.export_photos).grid(
            row=7, column=0, columnspan=2, sticky="ew", pady=(4, 2))
        if not CONVERTER_AVAILABLE:
            self.convert_btn.configure(state="disabled")

    # ----------------------------------------------------------------- main
    def _build_main(self, main):
        main.rowconfigure(1, weight=1)
        main.columnconfigure(0, weight=1)

        self.stats_var = tk.StringVar(
            value="No dump scanned yet - pick one in the left toolbar.")
        ttk.Label(main, textvariable=self.stats_var,
                  font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=4)

        nb = ttk.Notebook(main)
        nb.grid(row=1, column=0, sticky="nsew")
        self.notebook = nb

        # Titles tab
        f1 = ttk.Frame(nb)
        nb.add(f1, text="  Titles  ")
        self.titles_tree = SortableTree(f1, ("Title", "Total Play Time", "Days Played"),
                                        (320, 130, 100))
        self.titles_tree.pack_with_scroll(fill="both", expand=True)
        self.titles_tree.set_sort_key(
            "Total Play Time", lambda s: int(s.split(":")[0]) * 60 + int(s.split(":")[1]))
        self.titles_tree.set_sort_key("Days Played", int)

        # Days tab
        f2 = ttk.Frame(nb)
        nb.add(f2, text="  Days  ")
        self.days_tree = SortableTree(f2, ("Date", "Total Play Time", "Sessions",
                                           "'Others' Overflow"),
                                      (130, 130, 90, 130))
        self.days_tree.pack_with_scroll(fill="both", expand=True)
        self.days_tree.set_sort_key(
            "Total Play Time", lambda s: int(s.split(":")[0]) * 60 + int(s.split(":")[1]))
        self.days_tree.set_sort_key("Sessions", int)

        # Memos tab
        f3 = ttk.Frame(nb)
        nb.add(f3, text="  Memos  ")
        bar3 = ttk.Frame(f3)
        bar3.pack(fill="x", pady=2)
        ttk.Label(bar3, text="Filter:").pack(side="left")
        self.memo_filter = tk.StringVar()
        self.memo_filter.trace_add("write", lambda *_: self._refill_memos())
        ttk.Entry(bar3, textvariable=self.memo_filter, width=28).pack(side="left", padx=4)
        ttk.Label(bar3, text="(double-click a row to read the full message)",
                  foreground="#777").pack(side="left", padx=8)
        self.memos_tree = SortableTree(f3, ("Date", "Category", "Game / Channel",
                                            "Message", "Detail", "Photo"),
                                       (130, 110, 160, 180, 360, 60))
        self.memos_tree.pack_with_scroll(fill="both", expand=True)
        self.memos_tree.bind("<Double-Button-1>", self._on_memo_dblclick)

        # Photos tab: thumbnail gallery (left) + photo & its message (right)
        f4 = ttk.Frame(nb)
        nb.add(f4, text="  Photos  ")
        f4.columnconfigure(1, weight=1)
        f4.rowconfigure(0, weight=1)

        gal_frame = ttk.Frame(f4, width=400)
        gal_frame.grid(row=0, column=0, sticky="ns")
        gal_frame.grid_propagate(False)
        self.gallery_canvas = tk.Canvas(gal_frame, highlightthickness=0)
        gal_sb = ttk.Scrollbar(gal_frame, orient="vertical",
                               command=self.gallery_canvas.yview)
        self.gallery_inner = ttk.Frame(self.gallery_canvas)
        self.gallery_inner.bind(
            "<Configure>",
            lambda _e: self.gallery_canvas.configure(
                scrollregion=self.gallery_canvas.bbox("all")))
        self.gallery_canvas.create_window((0, 0), window=self.gallery_inner,
                                          anchor="nw")
        self.gallery_canvas.configure(yscrollcommand=gal_sb.set)
        self.gallery_canvas.pack(side="left", fill="both", expand=True)
        gal_sb.pack(side="right", fill="y")
        self.gallery_canvas.bind_all(
            "<MouseWheel>", lambda e: self.gallery_canvas.yview_scroll(
                -1 * (e.delta // 120), "units")
            if self.notebook.index("current") == 3 else None)

        det = ttk.Frame(f4)
        det.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        det.columnconfigure(0, weight=1)
        self.photo_img_lbl = ttk.Label(det, anchor="center")
        self.photo_img_lbl.grid(row=0, column=0, sticky="n", pady=4)
        pbtn = ttk.Frame(det)
        pbtn.grid(row=1, column=0, sticky="ew")
        self.save_photo_btn = ttk.Button(pbtn, text="Save photo as PNG…",
                                         command=self.save_photo,
                                         state="disabled")
        self.save_photo_btn.pack(side="left")
        self.open_msg_btn = ttk.Button(pbtn, text="Show in Memos tab",
                                       command=self.goto_photo_memo,
                                       state="disabled")
        self.open_msg_btn.pack(side="left", padx=6)
        self.photo_info = tk.Text(det, height=8, wrap="word", state="disabled")
        self.photo_info.grid(row=2, column=0, sticky="ew", pady=4)
        if PIL_AVAILABLE and AJPG_AVAILABLE:
            _hint = "No photos yet - scan a dump containing attached photos."
        elif not PIL_AVAILABLE:
            _hint = ("Photo viewing needs Pillow - open Tools → Dependencies… "
                     "to install it.")
        else:
            _hint = "Photo decoder unavailable - see the Log pane."
        self.photos_hint = tk.StringVar(value=_hint)
        ttk.Label(det, textvariable=self.photos_hint, foreground="#777",
                  wraplength=400, justify="left").grid(row=3, column=0,
                                                       sticky="w", pady=6)

        # Log pane (hidden by default - Toggle menu > Toggle Log)
        lf = ttk.LabelFrame(main, text="Log")
        lf.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        self.log_frame = lf
        self.log_text = tk.Text(lf, height=6, state="disabled", wrap="word")
        self.log_text.pack(fill="x")

    # ------------------------------------------------------------- helpers
    def toggle_sidebar(self):
        if self.sidebar.winfo_ismapped():
            self.sidebar.pack_forget()
            self.toggle_btn.configure(text="▶")
        else:
            self.sidebar.pack(side="left", fill="y", padx=(6, 0),
                              before=self.sidebar.master.winfo_children()[-1])
            self.toggle_btn.configure(text="◀")

    # ------------------------------------------------------------- menubar
    def _build_menubar(self):
        menubar = tk.Menu(self.root, tearoff=0)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="Select file… (Dump 1)",
                           command=self.input1.browse_file)
        m_file.add_command(label="Select path… (Dump 1)",
                           command=self.input1.browse_folder)
        m_file.add_separator()
        m_file.add_command(label="Select file… (Dump 2)",
                           command=self.input2.browse_file)
        m_file.add_command(label="Select path… (Dump 2)",
                           command=self.input2.browse_folder)
        m_file.add_separator()
        m_file.add_command(label="Exit Application", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=m_file)

        m_tools = tk.Menu(menubar, tearoff=0)
        m_tools.add_command(label="Export playtime .xlsx",
                            command=self.on_convert)
        m_tools.add_command(label="Export pictures to .png…",
                            command=self.export_photos)
        m_tools.add_command(label="Export All (.zip)",
                            command=self.on_export_all)
        m_tools.add_separator()
        m_tools.add_command(label="Clear cache…", command=self.clear_cache)
        m_tools.add_command(label="Dependencies…", command=self.show_deps)
        menubar.add_cascade(label="Tools", menu=m_tools)

        m_toggle = tk.Menu(menubar, tearoff=0)
        m_toggle.add_checkbutton(label="Toggle Log", variable=self.show_log_var,
                                 command=self._apply_log_visibility)
        m_toggle.add_checkbutton(label="Toggle Verbose Checklist",
                                 variable=self.verbose_var,
                                 command=self._apply_checklist_mode)
        m_toggle.add_checkbutton(label="Toggle Dark Theme",
                                 variable=self.dark_var,
                                 command=self._apply_theme)
        menubar.add_cascade(label="Toggle", menu=m_toggle)

        m_info = tk.Menu(menubar, tearoff=0)
        m_info.add_command(label="Logs", command=self.open_logs_folder)
        m_info.add_command(label="Create session log", command=self.show_logs)
        m_info.add_separator()
        m_info.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Info", menu=m_info)

        self.root.config(menu=menubar)

    def show_about(self):
        """Info → About: themed dialog (follows the light/dark setting)."""
        top = tk.Toplevel(self.root)
        top.title("About " + APP_TITLE)
        top.transient(self.root)
        top.resizable(False, False)
        if TTB_AVAILABLE and self.dark_var.get():
            top.configure(background="#222222")
        body = ttk.Frame(top, padding=18)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=APP_TITLE,
                  font=("", 13, "bold")).pack(anchor="w")
        ttk.Label(body, text=f"Version {APP_VERSION}"
                  ).pack(anchor="w", pady=(0, 10))
        ttk.Label(body, justify="left", wraplength=400, text=
            "Converts and makes viewable Wii Message Board dumps (playtime "
            "logs, memos, letters and attached aJPG photos) into "
            "spreadsheets and PNG images.").pack(anchor="w", pady=(0, 10))
        ttk.Label(body, justify="left", text=
            "Created by:\n  - Link4750 / ExoLink\n  - Kimi K3 (Moonshot AI)"
            ).pack(anchor="w", pady=(0, 10))
        ttk.Label(body, text="GNU General Public License v3.0\n https://github.com/ExoLink4750/WiiMessageBoard-Playtime-Photo-Converter"
                  ).pack(anchor="w", pady=(0, 12))
        ttk.Button(body, text="OK", style=_bs("primary"), width=10,
                   command=top.destroy).pack()

    def open_logs_folder(self):
        """Info → Logs: open the folder where diagnostic logs are written
        (created on demand)."""
        d = os.path.join(program_dir(), "logs")
        os.makedirs(d, exist_ok=True)
        self._open_folder(d)

    def clear_cache(self):
        """Tools menu: delete the cached scans and decoded photos."""
        if not os.path.isdir(CACHE_DIR):
            messagebox.showinfo(APP_TITLE, "The cache is already empty.")
            return
        total = 0
        for root, _dirs, files in os.walk(CACHE_DIR):
            total += sum(os.path.getsize(os.path.join(root, f))
                         for f in files)
        if not messagebox.askyesno(
                APP_TITLE,
                f"Delete the scan & photo cache ({total / 1048576:.1f} MB)?\n\n"
                f"{CACHE_DIR}\n\n"
                "The next scan of a dump will be fully re-parsed and its "
                "photos re-decoded."):
            return
        try:
            shutil.rmtree(CACHE_DIR)
        except Exception as e:
            messagebox.showerror(APP_TITLE,
                                 f"Could not clear the cache:\n{e}")
            return
        self.log(f"cache cleared: {CACHE_DIR}")
        messagebox.showinfo(APP_TITLE, "Cache cleared.")

    # ------------------------------------------------------------- log files
    @staticmethod
    def _program_dir():
        """Directory the running program lives in (script folder, or the
        folder containing the .exe for a PyInstaller build)."""
        if getattr(sys, "frozen", False):
            return os.path.dirname(os.path.abspath(sys.executable))
        return os.path.dirname(os.path.abspath(__file__))

    def write_log_file(self, reason="manual (Info → Logs)"):
        """Dump the session log plus environment info to
        <program dir>/logs/log(yyyy-mm-dd-hh-mm-ss).txt. Returns the path."""
        logs_dir = os.path.join(self._program_dir(), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        stamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        path = os.path.join(logs_dir, f"log({stamp}).txt")
        deps = ", ".join(f"{pip}={'ok' if dep_installed(mod) else 'MISSING'}"
                         for mod, pip, _req, _d in DEPENDENCIES)
        header = "\n".join([
            f"{APP_TITLE} - diagnostic log",
            f"program revision: {APP_VERSION}",
            f"written: {datetime.now():%Y-%m-%d %H:%M:%S}",
            f"reason: {reason}",
            f"python: {sys.version.split()[0]}   "
            f"frozen exe: {bool(getattr(sys, 'frozen', False))}",
            f"packages: {deps}",
            "converter core: " + ("ok" if CONVERTER_AVAILABLE
                                  else CONVERTER_IMPORT_ERROR),
            f"photo decoder: {AJPG_SOURCE}",
            f"inputs: {'; '.join(self._inputs()) or '(none)'}",
            "", "---------------- session log ----------------", ""])
        with open(path, "w", encoding="utf-8") as f:
            f.write(header + self.log_text.get("1.0", "end"))
        return path

    def show_logs(self):
        """Info → Logs: write the session log to disk on demand."""
        try:
            path = self.write_log_file()
        except Exception as e:
            messagebox.showerror(APP_TITLE,
                                 f"Could not write the log file:\n{e}")
            return
        if messagebox.askyesno(APP_TITLE, f"Log file written to:\n{path}\n\n"
                                          "Open the logs folder?"):
            self._open_folder(os.path.dirname(path))

    def _auto_log(self, reason, detail):
        """Something failed: dump a diagnostic log and tell the user where."""
        try:
            path = self.write_log_file(reason)
        except Exception as e:
            self.log(f"could not write log file: {e}")
            return
        messagebox.showwarning(APP_TITLE,
                               f"{detail}\n\nA diagnostic log was saved to:\n"
                               f"{path}")

    # ---------------------------------------------------- toggles + progress
    def _apply_theme(self):
        if not TTB_AVAILABLE:
            self.dark_var.set(False)
            messagebox.showinfo(
                APP_TITLE, "The optional 'ttkbootstrap' package is not "
                "installed.\nInstall it via Tools → Dependencies… to enable "
                "the themed UI with light/dark mode.")
            return
        # recreate the Style object: theme_use() alone leaves some widget
        # families (Treeview viewport, panes) stuck on the old theme
        self._ttb_style = ttkbootstrap.Style(
            theme="darkly" if self.dark_var.get() else "flatly")
        self._fix_tree_colors()
        self._restyle_plain_widgets()

    def _fix_tree_colors(self):
        # The empty area of a ttk.Treeview is drawn from 'fieldbackground',
        # which the themes leave unset (white). Force it to match.
        if not (TTB_AVAILABLE and self._ttb_style):
            return
        self._ttb_style.configure(
            "Treeview",
            fieldbackground=("#222222" if self.dark_var.get() else "white"))

    def _restyle_plain_widgets(self):
        """Match plain tk widgets (Text panes, menus, window chrome) to the
        current theme - ttkbootstrap only styles ttk widgets."""
        dark = TTB_AVAILABLE and self.dark_var.get()
        bg, fg = ("#2b2b2b", "#e8e8e8") if dark else ("white", "black")
        for w in (getattr(self, "log_text", None),
                  getattr(self, "photo_info", None)):
            if w is not None:
                w.configure(background=bg, foreground=fg, insertbackground=fg)
        win_bg = "#222222" if dark else ""
        mkw = (dict(background="#2b2b2b", foreground="#f0f0f0",
                    activebackground="#375a7f", activeforeground="#ffffff")
               if dark else
               dict(background="", foreground="",
                    activebackground="", activeforeground=""))
        def _walk_menu(m):
            m.configure(**mkw)
            end = m.index("end")
            if end is None:
                return
            for i in range(end + 1):
                if m.type(i) == "cascade":
                    _walk_menu(m.nametowidget(m.entrycget(i, "menu")))
        with contextlib.suppress(Exception):
            _walk_menu(self.root.nametowidget(self.root.cget("menu")))
            self.root.configure(background=win_bg)
            for w in self.root.winfo_children():
                if isinstance(w, tk.Toplevel):
                    w.configure(background=win_bg)

    def _apply_log_visibility(self):
        if self.show_log_var.get():
            self.log_frame.grid()
        else:
            self.log_frame.grid_remove()

    def _apply_checklist_mode(self):
        if not self._check_rows:
            return
        verbose = self.verbose_var.get()
        for key, (simple, verbose_text, ok) in self._check_rows.items():
            self.check_vars[key].set(verbose_text if verbose else simple)
            color = "#777" if ok is None else ("#087008" if ok else "#a33")
            self.check_labels[key].configure(foreground=color)

    def progress_set(self, value, maximum, text):
        """Determinate progress with the status text inside the bar."""
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=max(1, maximum),
                                value=value)
        self.progress_text.configure(text=text)

    def progress_indeterminate(self, text):
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self.progress_text.configure(text=text)

    def progress_clear(self):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.progress_text.configure(text="")

    def browse_out(self):
        path = filedialog.askdirectory(title="Select workbook output folder")
        if path:
            self.out_dir.set(path)

    def browse_photo_out(self):
        path = filedialog.askdirectory(title="Select photo output folder")
        if path:
            self.photo_out_dir.set(path)

    @staticmethod
    def _expand_out_name(name):
        """Fill the (yyyy-mm-dd-hh-mm-ss) template token at export time."""
        return name.replace(TS_TOKEN,
                            datetime.now().strftime("%Y-%m-%d-%H-%M-%S"))

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _inputs(self):
        return [p for p in (self.input1.get(), self.input2.get()) if p]

    # ---------------------------------------------------------- dependencies
    def show_deps(self):
        DepsDialog(self)

    def _startup_dep_check(self):
        if not CONVERTER_AVAILABLE:
            self.log("converter core failed to load: " + CONVERTER_IMPORT_ERROR)
        if not PIL_AVAILABLE:
            self.log("Pillow is not available: " + PIL_IMPORT_ERROR)
        if not AJPG_AVAILABLE:
            self.log("AJPG photo decoder unavailable: " + AJPG_IMPORT_ERROR)
        elif AJPG_SOURCE == "embedded":
            self.log("ajpg_to_png.py not found - using the built-in decoder copy")
        missing_req = [p for m, p, req, _ in DEPENDENCIES
                       if req and not dep_installed(m)]
        if missing_req or not CONVERTER_AVAILABLE:
            self.log("missing required package(s): "
                     + (", ".join(missing_req) or "see above")
                     + " - opening the dependency manager…")
            DepsDialog(self)
        else:
            missing_opt = [p for m, p, req, _ in DEPENDENCIES
                           if not req and not dep_installed(m)]
            if missing_opt:
                self.log("optional package(s) not installed: "
                         + ", ".join(missing_opt) + " (see Dependencies…)")
            else:
                self.log("all dependencies installed")

    def install_dep(self, pipname, on_done=None):
        if getattr(sys, "frozen", False):
            # a packaged exe has no pip of its own: use a system Python to
            # install into 'python_packages' next to the exe (on sys.path)
            py = next((c for c in (["py", "-3"], ["python"], ["python3"])
                       if shutil.which(c[0])), None)
            if not py:
                messagebox.showinfo(
                    APP_TITLE,
                    "No system Python found to fetch the package with.\n"
                    "Either install Python (python.org), or rebuild with the "
                    "updated Build Executable.cmd so it is bundled.")
                return
            cmd = py + ["-m", "pip", "install", "--upgrade",
                        "--target", _EXTRA_PKG_DIR, pipname]
        else:
            cmd = [sys.executable, "-m", "pip", "install", pipname]
        self.log("running: " + " ".join(cmd))
        self.progress_indeterminate(f"Installing {pipname}…")
        self.status.set(f"Installing {pipname}…")

        def work():
            try:
                proc = subprocess.run(cmd, capture_output=True, text=True,
                                      timeout=600)
                tail = "\n".join((proc.stdout + "\n" + proc.stderr)
                                 .strip().splitlines()[-6:])
                if tail:
                    self.queue.put(("log", "[pip] "
                                    + tail.replace("\n", "\n[pip] ")))
                self.queue.put(("dep_installed", pipname,
                                proc.returncode == 0, on_done))
            except Exception as e:
                self.queue.put(("log", f"install of {pipname} failed: {e}"))
                self.queue.put(("dep_installed", pipname, False, on_done))

        threading.Thread(target=work, daemon=True).start()

    def _enable_pil(self):
        """Import Pillow after an install, without restarting the app."""
        global Image, ImageTk, PIL_AVAILABLE, PIL_IMPORT_ERROR
        try:
            _augment_sys_path()
            importlib.invalidate_caches()
            Image = importlib.import_module("PIL.Image")
            ImageTk = importlib.import_module("PIL.ImageTk")
            PIL_AVAILABLE, PIL_IMPORT_ERROR = True, ""
            return True
        except Exception as e:
            PIL_IMPORT_ERROR = f"{type(e).__name__}: {e}"
            return False

    def _enable_converter(self):
        """Import the converter core after installing openpyxl."""
        global CONVERTER_AVAILABLE, CONVERTER_IMPORT_ERROR
        try:
            _augment_sys_path()
            importlib.invalidate_caches()
            mod = importlib.import_module("playtime_log_to_excel")
            importlib.reload(mod)
            globals().update({n: getattr(mod, n) for n in (
                "TITLE_IDS", "VFF_PATH_RE", "SourceBundle", "build_workbook",
                "dedup_bundles", "is_letter", "load_source", "memo_category",
                "memo_game_detail", "open_source", "parse_board_date",
                "parse_cdbfile", "parse_memo_text", "readable_title",
                "title_id_name", "vff_record_date")})
            CONVERTER_AVAILABLE, CONVERTER_IMPORT_ERROR = True, ""
            self.convert_btn.configure(state="normal")
            return True
        except Exception as e:
            CONVERTER_IMPORT_ERROR = f"{type(e).__name__}: {e}"
            return False

    def _enable_ttb(self):
        """Import ttkbootstrap after a 1-click install, no restart needed."""
        global ttkbootstrap, TTB_AVAILABLE
        try:
            _augment_sys_path()
            importlib.invalidate_caches()
            ttkbootstrap = importlib.import_module("ttkbootstrap")
            TTB_AVAILABLE = True
            self._ttb_style = ttkbootstrap.Style(
                theme="darkly" if self.dark_var.get() else "flatly")
            self._fix_tree_colors()
            self._restyle_plain_widgets()
            return True
        except Exception as e:
            self.log(f"ttkbootstrap installed but could not load: {e}")
            return False

    def _after_dep_install(self, pipname, ok, on_done):
        if ok:
            self.log(f"{pipname} installed successfully")
            if pipname == "pillow" and self._enable_pil():
                self.log("photo support enabled")
            elif pipname == "openpyxl":
                self._enable_converter()
            elif pipname == "tkinterdnd2":
                self.log("restart the app to enable drag && drop")
            elif pipname == "ttkbootstrap":
                if self._enable_ttb():
                    self.log("themed UI enabled (see Toggle → Dark Theme)")
            # rarfile is picked up automatically on the next scan
            if self._inputs() and CONVERTER_AVAILABLE:
                self.log("re-scanning with the new package…")
                self.start_scan()
        else:
            self.log(f"could not install {pipname} - try running "
                     f"'{sys.executable} -m pip install {pipname}' "
                     "in a terminal")
        self.status.set("Ready")
        self.progress_clear()
        if callable(on_done):
            with contextlib.suppress(Exception):
                on_done()

    # ------------------------------------------------------------------ scan
    def schedule_scan(self):
        if self._scan_job:
            self.root.after_cancel(self._scan_job)
        self._scan_job = self.root.after(500, self.start_scan)

    def start_scan(self):
        if not CONVERTER_AVAILABLE:
            self.scan_status.set("Converter core unavailable - see Dependencies…")
            return
        paths = self._inputs()
        self.bundles, self.bundles_key = [], None
        self._set_checklist(None)
        self._memo_photo_count = {}
        self._photo_by_key = {}
        if hasattr(self, "gallery_inner"):
            self._populate_photos([])
        if not paths:
            self.scan_status.set("Select dump(s), then press Scan Dumps.")
            return
        self._scan_gen += 1
        gen = self._scan_gen
        self.scan_status.set("Scanning…")
        self.status.set("Scanning…")
        self.progress_set(0, 2 * len(paths), "Scanning…")
        threading.Thread(target=self._scan_worker, args=(gen, tuple(paths)),
                         daemon=True).start()

    def _scan_worker(self, gen, paths):
        qlog = lambda m: self.queue.put(("log", m))
        qlog(f"scan started ({len(paths)} input(s))")
        bundles, errors = [], []
        keys = {}   # input path -> cache key (None = do not cache)
        for i, p in enumerate(paths):
            self.queue.put(("progress", gen, i, 2 * len(paths),
                            f"Reading {os.path.basename(p)}…"))
            qlog(f"reading: {p}")
            try:
                key = _dump_cache_key(p)
                keys[p] = key
                b = _cache_read("scan", key) if key else _CACHE_MISS
                if b is _CACHE_MISS:
                    b = load_source(p)
                    if key:
                        _cache_write("scan", key, b)
                else:
                    qlog(f"scan cache hit: {os.path.basename(p)}")
                bundles.append(b)
                qlog(f"parsed {b.label}: {b.kind}, "
                     f"{len(b.playlogs)} play-time log(s), {len(b.memos)} "
                     f"memo(s), {b.photos} photo(s), {b.arcs} archive(s)")
            except Exception as e:
                errors.append(f"{os.path.basename(p)}: {e}")
                qlog(f"ERROR reading {p}: {e}")
        photos = []
        if not errors and AJPG_AVAILABLE:
            for j, b in enumerate(bundles):
                self.queue.put(("progress", gen, len(paths) + j,
                                2 * len(paths),
                                "Scanning for attached photos…"))
                try:
                    key = keys.get(b.path)
                    found = _cache_read("photoscan", key) if key \
                        else _CACHE_MISS
                    if found is _CACHE_MISS:
                        found = collect_photos(b)
                        if key:
                            _cache_write("photoscan", key, found)
                    else:
                        qlog(f"photo-scan cache hit: {b.label}")
                    photos.extend(found)
                    qlog(f"photo scan {b.label}: {len(found)} attached "
                         "photo(s)")
                except Exception as e:
                    errors.append(f"{b.label}: photo scan: {e}")
                    qlog(f"ERROR scanning photos in {b.label}: {e}")
            before = len(photos)
            photos = dedup_photos(photos)
            qlog(f"photo scan complete: {len(photos)} unique photo(s) "
                 f"({before - len(photos)} duplicate copy(ies) merged)")
        self.queue.put(("scan", gen, paths, bundles, errors, photos))
        if photos and not errors:
            self._decode_worker(gen, photos)

    def _decode_worker(self, gen, photos):
        qlog = lambda m: self.queue.put(("log", m))
        if not (PIL_AVAILABLE and AJPG_AVAILABLE):
            qlog("photo decode skipped: decoder and/or Pillow unavailable")
            self.queue.put(("photos", gen, [], []))
            return
        qlog(f"decoding {len(photos)} photo(s)…")
        ok, fails, hits = [], [], 0
        for i, it in enumerate(photos):
            try:
                it.image, cached = decode_photo_image(it)
                hits += cached
                ok.append(it)
            except Exception as e:
                fails.append(it)
                qlog(f"photo decode failed ({it.source}): {e}")
            self.queue.put(("photos_progress", gen, i + 1, len(photos)))
        qlog(f"decode finished: {len(ok)}/{len(photos)} ok"
             + (f", {len(fails)} FAILED" if fails else "")
             + (f" ({hits} from cache)" if hits else ""))
        self.queue.put(("photos", gen, ok, fails))

    def _set_checklist(self, bundles):
        if not bundles:
            self._check_rows = {key: ("—  " + CHECK_SIMPLE[key][0],
                                      "—  " + text, None)
                                for key, text in CHECK_ITEMS}
            self._apply_checklist_mode()
            return
        eff_logs, eff_memos, dup_l, dup_m = dedup_bundles(bundles)
        self._eff_logs, self._eff_memos = eff_logs, eff_memos
        cats = {}
        for m in eff_memos:
            cats[memo_category(m)] = cats.get(memo_category(m), 0) + 1
        letters = cats.get("letter", 0) + cats.get("channel", 0)
        tids = sorted({m.title_id for m in eff_memos if m.title_id})
        photos = max((b.photos for b in bundles), default=0)
        arcs = max((b.arcs for b in bundles), default=0)
        self._dedup_note = (f" ({dup_l} log(s) and {dup_m} memo(s) exist in both "
                            "dumps and will be counted once)"
                            if (dup_l or dup_m) else "")

        counts = {"logs": len({l.day for l in eff_logs}),
                  "achv": cats.get("achievement", 0),
                  "letters": letters,
                  "personal": cats.get("personal", 0),
                  "titleids": len(tids),
                  "photos": photos,
                  "arcs": arcs}
        suffix = {"logs": f" days, {sum(len(l.entries) for l in eff_logs)} entries",
                  "achv": " memos", "letters": " letters", "personal": " notes",
                  "titleids": ": " + ", ".join(tids[:6]) + ("…" if len(tids) > 6 else ""),
                  "photos": " photos", "arcs": " archives"}
        self._check_rows = {}
        for key, text in CHECK_ITEMS:
            n = counts[key]
            label, unit = CHECK_SIMPLE[key]
            if n:
                self._check_rows[key] = (
                    f"✓  {label}: {n} {unit}",
                    f"✓  {text}   [{n}{suffix[key]}]", True)
            else:
                self._check_rows[key] = (f"✗  {label}", f"✗  {text}", False)
        self._apply_checklist_mode()

        # main-area tables
        total_min = sum(l.total for l in eff_logs)
        self.stats_var.set(
            f"{len({l.day for l in eff_logs})} days  |  "
            f"total play time {fmt_hhmm(total_min)}  |  "
            f"{len(eff_memos)} memos  |  "
            f"{len({t for l in eff_logs for t, _ in l.entries})} titles")

        tstats = {}
        for l in eff_logs:
            for t, mins in l.entries:
                m = tstats.setdefault(t, [0, set()])
                m[0] += mins
                m[1].add(l.day)
        self.titles_tree.refill(sorted(
            ((t, fmt_hhmm(v[0]), str(len(v[1]))) for t, v in tstats.items()),
            key=lambda r: -(int(r[1].split(":")[0]) * 60 + int(r[1].split(":")[1]))))

        dstats = {}
        for l in eff_logs:
            d = dstats.setdefault(l.day, [0, 0, False])
            d[0] += l.total
            d[1] += len(l.entries)
            d[2] = d[2] or any(t == "Others" for t, _ in l.entries)
        self.days_tree.refill(sorted(
            (d, fmt_hhmm(v[0]), str(v[1]), "Yes" if v[2] else "No")
            for d, v in dstats.items()))

        self._build_memo_rows()

    def _build_memo_rows(self):
        """(Re)build the Memos tab rows; includes the 📷 attachment marker."""
        self._memo_rows = []
        self._memo_keys = []
        self._memo_by_source = {m.source: m for m in self._eff_memos}
        for m in self._eff_memos:
            game, detail = memo_game_detail(m)
            n = self._memo_photo_count.get(m.source, m.photos or 0)
            self._memo_rows.append((
                m.date.strftime("%Y-%m-%d %H:%M"),
                memo_category(m).capitalize(),
                game, readable_title(m.title), detail,
                "📷" * n if n else ""))
            self._memo_keys.append(m.source)
        pairs = sorted(zip(self._memo_rows, self._memo_keys),
                       key=lambda p: p[0])
        self._memo_rows = [p[0] for p in pairs]
        self._memo_keys = [p[1] for p in pairs]
        self._refill_memos()

    def _refill_memos(self):
        if not hasattr(self, "_memo_rows"):
            return
        f = self.memo_filter.get().strip().lower()
        pairs = [(r, k) for r, k in zip(self._memo_rows, self._memo_keys)
                 if not f or any(f in str(c).lower() for c in r)]
        self.memos_tree.delete(*self.memos_tree.get_children(""))
        self._memo_shown = pairs
        for i, (row, _key) in enumerate(pairs):
            self.memos_tree.insert("", "end", iid=str(i), values=row)

    # ------------------------------------------------------------- photos
    def _on_memo_dblclick(self, _event):
        sel = self.memos_tree.selection()
        if not sel:
            return
        key = self._memo_shown[int(sel[0])][1]
        memo = getattr(self, "_memo_by_source", {}).get(key)
        if memo is not None:
            self.show_memo(memo, key)
        else:  # fallback: old behaviour, jump straight to the photo
            photo = self._photo_by_key.get(key)
            if photo and photo.image is not None:
                self.notebook.select(3)
                self.select_photo(photo)

    def show_memo(self, memo, key=None):
        """Popup showing the COMPLETE text of a board message/letter (the
        Memos tab 'Detail' column only shows a short preview)."""
        top = tk.Toplevel(self.root)
        top.title("Message - " + memo.date.strftime("%Y-%m-%d %H:%M"))
        if TTB_AVAILABLE and self.dark_var.get():
            top.configure(background="#222222")
        top.minsize(480, 320)
        top.geometry("560x440")

        head = ttk.Frame(top, padding=(10, 8, 10, 0))
        head.pack(fill="x")
        cat = memo_category(memo).capitalize()
        game = (title_id_name(memo.title_id) if memo.title_id
                and memo.title_id != "HAEA" else cat)
        where = game if game == cat else f"{game}   |   {cat}"
        ttk.Label(head, text=f"{memo.date:%Y-%m-%d %H:%M}   |   {where}",
                  font=("", 10, "bold")).pack(anchor="w")
        ttk.Label(head, text=memo.source, foreground="#777", wraplength=520,
                  justify="left").pack(anchor="w", pady=(2, 0))

        # attached photo(s) shown inline, above the message text
        photos = [it for it in self.photo_items
                  if it.key == (key or memo.source) and it.image is not None]
        top._img_refs = []
        if photos:
            pf = ttk.Frame(top, padding=(10, 6, 10, 0))
            pf.pack(fill="x")
            for it in photos[:4]:
                img = it.image.copy()
                img.thumbnail((250, 190))
                ref = ImageTk.PhotoImage(img)
                top._img_refs.append(ref)   # keep alive
                ttk.Label(pf, image=ref).pack(side="left", padx=4)

        txt_frame = ttk.Frame(top)
        txt_frame.pack(fill="both", expand=True, padx=10, pady=8)
        body = tk.Text(txt_frame, wrap="word", padx=6, pady=6)
        if TTB_AVAILABLE and self.dark_var.get():
            body.configure(background="#2b2b2b", foreground="#e8e8e8",
                           insertbackground="#e8e8e8")
        sb = ttk.Scrollbar(txt_frame, orient="vertical", command=body.yview)
        body.configure(yscrollcommand=sb.set)
        body.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        title = readable_title(memo.title)
        body.insert("1.0", (title + "\n\n" if title else "")
                    + (memo.flat or "(no text)"))
        body.configure(state="disabled")

        btns = ttk.Frame(top, padding=(10, 0, 10, 8))
        btns.pack(fill="x")
        photo = self._photo_by_key.get(key if key is not None else memo.source)
        if photo is not None and photo.image is not None:
            ttk.Button(btns, text="Show attached photo",
                       command=lambda: (top.destroy(),
                                        self.notebook.select(3),
                                        self.select_photo(photo))
                       ).pack(side="left")
        ttk.Button(btns, text="Copy text",
                   command=lambda: (top.clipboard_clear(),
                                    top.clipboard_append(memo.flat))
                   ).pack(side="left", padx=6)
        ttk.Button(btns, text="Close", command=top.destroy).pack(side="right")

    def select_photo(self, item):
        """Show a photo in the detail pane alongside its message."""
        self._sel_photo = item
        if item.image is None:
            return
        img = item.image
        scale = min(1.0, 500 / img.width, 320 / img.height)
        if scale < 1.0:
            img = img.resize((max(1, int(img.width * scale)),
                              max(1, int(img.height * scale))))
        self._detail_img_ref = ImageTk.PhotoImage(img)
        self.photo_img_lbl.configure(image=self._detail_img_ref)
        game = title_id_name(item.title_id) if item.title_id else \
            "Wii Message Board"
        info = (f"Date:    {item.date_str}\n"
                f"From:    {game}\n"
                f"Message: {item.msg_title or '(no title)'}\n"
                f"{item.msg_body or ''}\n"
                f"Source:  {item.source}  ({item.dump})")
        self.photo_info.configure(state="normal")
        self.photo_info.delete("1.0", "end")
        self.photo_info.insert("1.0", info)
        self.photo_info.configure(state="disabled")
        self.save_photo_btn.configure(state="normal")
        self.open_msg_btn.configure(state="normal")
        for btn, it in self._gallery_btns:
            btn.configure(style=("Sel.TButton" if it is item
                                 else "TButton"))

    def goto_photo_memo(self):
        """Jump from the photo detail pane to its row in the Memos tab."""
        item = self._sel_photo
        if not item:
            return
        self.notebook.select(2)
        self.memo_filter.set("")
        for i, (_row, key) in enumerate(getattr(self, "_memo_shown", [])):
            if key == item.key:
                iid = str(i)
                self.memos_tree.selection_set(iid)
                self.memos_tree.see(iid)
                break

    def save_photo(self):
        item = self._sel_photo
        if not item or item.image is None:
            return
        base = f"wii photo {item.date_str.replace(':', '-')}" \
            if item.date else "wii photo"
        path = filedialog.asksaveasfilename(
            title="Save photo", initialfile=base + ".png",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png"), ("JPEG", "*.jpg")])
        if path:
            item.image.save(path)
            self.log(f"photo saved to: {path}")

    def _populate_photos(self, photos):
        self.photo_items = photos
        self._memo_photo_count = {}
        self._photo_by_key = {}
        for it in photos:
            self._memo_photo_count[it.key] = \
                self._memo_photo_count.get(it.key, 0) + 1
            self._photo_by_key.setdefault(it.key, it)
        self._build_memo_rows()
        # thumbnail gallery
        for w in self.gallery_inner.winfo_children():
            w.destroy()
        self._thumb_refs = []
        self._gallery_btns = []
        if not photos:
            pend = getattr(self, "_pending_photos", [])
            if pend and not (PIL_AVAILABLE and AJPG_AVAILABLE):
                self.photos_hint.set(
                    f"{len(pend)} attached photo(s) found, but they cannot be "
                    "decoded because Pillow is not available - open Tools → "
                    "Dependencies… to install it.")
            elif PIL_AVAILABLE and AJPG_AVAILABLE:
                self.photos_hint.set("No attached photos found in the scanned "
                                     "dump(s).")
            return
        for i, it in enumerate(photos):
            thumb = it.image.copy()
            thumb.thumbnail((170, 130))
            tkimg = ImageTk.PhotoImage(thumb)
            self._thumb_refs.append(tkimg)
            btn = ttk.Button(self.gallery_inner, image=tkimg,
                             text=f"{it.date_str}\n{it.msg_title[:24]}",
                             compound="top",
                             command=lambda x=it: self.select_photo(x))
            btn.grid(row=i // 2, column=i % 2, padx=4, pady=4, sticky="n")
            self._gallery_btns.append((btn, it))
        self.photos_hint.set(
            f"{len(photos)} photo(s) decoded - click one to view it with its "
            "message.")
        self.select_photo(photos[0])

    # --------------------------------------------------------------- exports
    @staticmethod
    def _photo_png_name(it, idx):
        date = (it.date_str.replace(":", "-") if it.date
                else f"unknown date {idx:03d}")
        game = title_id_name(it.title_id) if it.title_id \
            else "Wii Message Board"
        game = re.sub(r'[\\/:*?"<>|]', "", game).strip() or "Wii Message Board"
        dump = re.sub(r'[\\/:*?"<>|]', "", it.dump).strip() or "dump"
        return f"[{date}][{game}][{dump}].png"

    @staticmethod
    def _unique_name(name, used):
        if name not in used:
            used.add(name)
            return name
        base, n = name[:-4], 1
        while name in used:
            n += 1
            name = f"{base} ({n}).png"
        used.add(name)
        return name

    def export_photos(self):
        """Tools menu: write every attached photo to PNG files in a folder."""
        photos = self.photo_items or getattr(self, "_pending_photos", [])
        if not photos:
            messagebox.showinfo(APP_TITLE, "No photos to export - scan a dump "
                                           "containing attached photos first.")
            return
        if not (PIL_AVAILABLE and AJPG_AVAILABLE):
            messagebox.showwarning(APP_TITLE, "Photo export needs Pillow - see "
                                              "Tools → Dependencies…")
            return
        out_dir = self.photo_out_dir.get().strip() or DEFAULT_PHOTOS_DIR
        self._busy = True
        self.progress_indeterminate("Exporting photos…")
        threading.Thread(target=self._export_photos_worker,
                         args=(list(photos), out_dir), daemon=True).start()

    def _export_photos_worker(self, photos, out_dir):
        try:
            os.makedirs(out_dir, exist_ok=True)
            used = set()
            for i, it in enumerate(photos):
                try:
                    if it.image is None:
                        it.image, _cached = decode_photo_image(it)
                except Exception as e:
                    self.queue.put(("log",
                                    f"photo skipped ({it.source}): {e}"))
                    continue
                name = self._unique_name(self._photo_png_name(it, i + 1), used)
                it.image.save(os.path.join(out_dir, name))
                self.queue.put(("xprogress", i + 1, len(photos),
                                f"Exporting photos… {i + 1}/{len(photos)}"))
            self.queue.put(("export_done",
                            f"{len(used)} photo(s) saved to {out_dir}",
                            out_dir))
        except Exception:
            self.queue.put(("error", traceback.format_exc()))

    def on_export_all(self):
        """Tools menu: workbook + all photos packed into one .zip file."""
        if self._busy:
            return
        if not CONVERTER_AVAILABLE:
            messagebox.showwarning(
                APP_TITLE, "The converter core is not available:\n"
                + CONVERTER_IMPORT_ERROR + "\n\nSee Tools → Dependencies…")
            return
        paths = self._inputs()
        if not paths:
            messagebox.showwarning(APP_TITLE, "Please select at least one input dump.")
            return
        out_dir = self.out_dir.get().strip() or DEFAULT_WORKBOOK_DIR
        out_name = self._expand_out_name(
            self.out_name.get().strip() or DEFAULT_OUT_NAME)
        zip_name = (out_name[:-5] if out_name.lower().endswith(".xlsx")
                    else out_name) + ".zip"
        key = tuple(paths)
        bundles = (self.bundles if self.bundles_key == key
                   and len(self.bundles) == len(paths) else None)
        photos = list(self.photo_items or getattr(self, "_pending_photos", []))
        self._busy = True
        self.convert_btn.configure(state="disabled")
        self.progress_indeterminate("Exporting workbook + photos…")
        self.status.set("Exporting…")
        threading.Thread(
            target=self._export_all_worker,
            args=(paths, bundles, out_dir, out_name, zip_name, photos),
            daemon=True).start()

    def _export_all_worker(self, paths, bundles, out_dir, out_name, zip_name,
                           photos):
        try:
            if bundles is None:
                bundles = [load_source_cached(p) for p in paths]
            tmpdir = tempfile.mkdtemp(prefix="wmbc_")
            xlsx = convert(bundles, tmpdir, out_name,
                           log=lambda m: self.queue.put(("log", m)))
            zpath = os.path.join(out_dir, zip_name)
            n = 0
            with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(xlsx, os.path.basename(xlsx))
                if photos and PIL_AVAILABLE and AJPG_AVAILABLE:
                    used = set()
                    for i, it in enumerate(photos):
                        try:
                            if it.image is None:
                                it.image, _cached = decode_photo_image(it)
                        except Exception as e:
                            self.queue.put(("log",
                                            f"photo skipped ({it.source}): {e}"))
                            continue
                        name = self._unique_name(
                            self._photo_png_name(it, i + 1), used)
                        p = os.path.join(tmpdir, name)
                        it.image.save(p)
                        z.write(p, "photos/" + name)
                        n += 1
                        self.queue.put(("xprogress", i + 1, len(photos),
                                        f"Packing photos… {i + 1}/{len(photos)}"))
                elif photos:
                    self.queue.put(("log", "photos skipped: Pillow unavailable "
                                           "(see Tools → Dependencies…)"))
            self.queue.put(("export_done",
                            f"Exported {os.path.basename(xlsx)} + {n} photo(s) "
                            f"to {zpath}", out_dir))
        except Exception:
            self.queue.put(("error", traceback.format_exc()))

    # --------------------------------------------------------------- convert
    def on_convert(self):
        if self._busy:
            return
        if not CONVERTER_AVAILABLE:
            messagebox.showwarning(
                APP_TITLE, "The converter core is not available:\n"
                + CONVERTER_IMPORT_ERROR + "\n\nSee Dependencies…")
            return
        paths = self._inputs()
        if not paths:
            messagebox.showwarning(APP_TITLE, "Please select at least one input dump.")
            return
        out_dir = self.out_dir.get().strip() or DEFAULT_WORKBOOK_DIR
        out_name = self._expand_out_name(
            self.out_name.get().strip() or DEFAULT_OUT_NAME)
        key = tuple(paths)
        bundles = (self.bundles if self.bundles_key == key
                   and len(self.bundles) == len(paths) else None)
        self._busy = True
        self.convert_btn.configure(state="disabled")
        self.progress_indeterminate("Converting to spreadsheet…")
        self.status.set("Converting…")
        threading.Thread(target=self._convert_worker,
                         args=(paths, bundles, out_dir, out_name), daemon=True).start()

    def _convert_worker(self, paths, bundles, out_dir, out_name):
        try:
            if bundles is None:
                bundles = [load_source_cached(p) for p in paths]
            out = convert(bundles, out_dir, out_name,
                          log=lambda m: self.queue.put(("log", m)))
            self.queue.put(("done", out))
        except Exception:
            self.queue.put(("error", traceback.format_exc()))

    # ---------------------------------------------------------------- queue
    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]
                if kind == "scan":
                    _k, gen, paths, bundles, errors, photos = item
                    if gen != self._scan_gen:
                        continue
                    self.bundles = bundles if not errors else []
                    self.bundles_key = tuple(paths) if not errors else None
                    self._set_checklist(bundles)
                    if errors:
                        self.scan_status.set("Could not read: " + "; ".join(errors))
                        self.status.set("Scan failed")
                        self.progress_clear()
                        self._auto_log("scan error",
                                       "Some input could not be read - see "
                                       "the log pane for details.")
                    else:
                        eff_l = len(self._eff_logs)
                        eff_m = len(self._eff_memos)
                        note = getattr(self, "_dedup_note", "")
                        self.scan_status.set(
                            f"Found {eff_l} play-time log(s) and {eff_m} memo(s).{note}")
                        self.status.set("Ready")
                    self._pending_photos = photos
                    if not photos:
                        self._populate_photos([])
                        if not errors:
                            self.progress_clear()
                    elif PIL_AVAILABLE and AJPG_AVAILABLE:
                        self.status.set(f"Decoding 0/{len(photos)} photos…")
                        self.progress_set(0, len(photos),
                                          f"Decoding photos… 0/{len(photos)}")
                        self.photos_hint.set(
                            f"Decoding {len(photos)} attached photo(s)…")
                    else:
                        self.progress_clear()
                elif kind == "progress":
                    _k, gen, val, mx, text = item
                    if gen == self._scan_gen:
                        self.progress_set(val, mx, text)
                elif kind == "xprogress":
                    self.progress_set(item[1], item[2], item[3])
                elif kind == "photos_progress":
                    _k, gen, done, total = item
                    if gen == self._scan_gen:
                        self.status.set(f"Decoding {done}/{total} photos…")
                        self.progress_set(done, total,
                                          f"Decoding photos… {done}/{total}")
                elif kind == "photos":
                    _k, gen, photos, fails = item
                    if gen != self._scan_gen:
                        continue
                    self._populate_photos(photos)
                    self.status.set("Ready")
                    self.progress_clear()
                    if fails:
                        self._auto_log("photo decode error",
                                       f"{len(fails)} attached photo(s) "
                                       "could not be decoded.")
                elif kind == "dep_installed":
                    self._after_dep_install(item[1], item[2], item[3])
                elif kind == "log":
                    self.log(item[1])
                elif kind == "done":
                    self._busy = False
                    self.convert_btn.configure(state="normal")
                    self.status.set(f"Done: {item[1]}")
                    self.progress_clear()
                    self.log(f"Output written to: {item[1]}")
                    if messagebox.askyesno(APP_TITLE, "Conversion finished.\nOpen the output folder?"):
                        self._open_folder(os.path.dirname(item[1]))
                elif kind == "export_done":
                    self._busy = False
                    self.convert_btn.configure(state="normal")
                    self.status.set("Ready")
                    self.progress_clear()
                    self.log(item[1])
                    if messagebox.askyesno(APP_TITLE,
                                           item[1] + "\n\nOpen the output folder?"):
                        self._open_folder(item[2])
                elif kind == "error":
                    self._busy = False
                    self.convert_btn.configure(state="normal")
                    self.status.set("Error")
                    self.progress_clear()
                    self.show_log_var.set(True)
                    self._apply_log_visibility()
                    self.log(item[1])
                    self._auto_log("conversion/export error",
                                   "The operation failed - the traceback is "
                                   "in the log pane.")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    @staticmethod
    def _open_folder(path):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa: S606
            elif sys.platform == "darwin":
                os.system(f'open "{path}"')
            else:
                os.system(f'xdg-open "{path}"')
        except Exception:
            pass


def main():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
