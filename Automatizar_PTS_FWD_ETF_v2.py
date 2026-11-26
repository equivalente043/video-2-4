# -*- coding: utf-8 -*-
"""PTS FWD & ETF | Excel + Bloomberg | automatización histórica.
Windows / Excel de escritorio / Bloomberg abierto / pywin32.
V2: conserva la sesión activa y el complemento Bloomberg; no cambia seguridad de macros.
No modifica el MAIN ni reemplaza reportes existentes.
La validación de referencia se basa en los XLSX del 17 y 18 de noviembre de 2026.
"""
import argparse, base64, csv, datetime as dt, hashlib, json, logging, math, os
from pathlib import Path
import posixpath, re, shutil, sys, tempfile, time, uuid, zipfile, zlib
import xml.etree.ElementTree as ET

MAIN_DIR = Path(r'C:\Users\p669213\OneDrive - Principal Financial Group\Maria')
OUTPUT_ROOT = Path(r'C:\Users\p669213\Principal Financial Group\Gerencia Finanzas AFP - Control y Cumplimiento Inversiones\002 Cumplimiento\Conflictos de Interés\Control de precios de transacciones\Pts FWD & ETF - Bloomberg')
DEFAULT_START, DEFAULT_END = '20260904', '20261117'
EPOCH = dt.date(1899, 12, 30)
N = {'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
ERRORS = {2000:'#NULL!',2007:'#DIV/0!',2015:'#VALUE!',2023:'#REF!',2029:'#NAME?',2036:'#NUM!',2042:'#N/A',2043:'#GETTING_DATA',2045:'#SPILL!',2050:'#CALC!'}
COM_ERRORS = {(-2146828288+k):v for k,v in ERRORS.items()}
LOG = logging.getLogger('PTS')

# Referencias de validación integradas (extraídas de los archivos entregados).
BASELINE_B85 = (
    'c-ri}>v9~)ktX^oJbnHOV8%T{{|Z%^Sw#{+H9)Y*mdTFwXk@#xoss59n%y01<GZi>=N^%XOt1lxAp5cy*|3q7adr3bdwBl%&0l}`'
    '_doyRo8SN8oAb@#;+x-nv)NspZ2F%5>R%7^@7d+WB|mM?`QPjBHpgpzdNpic9p#Vzz1^I1^P8*v%N>3?PCwmT^G}<bo%}JL9p#UU'
    '@AVh@S4N^gHdoj3$FaYVKl)K$eZRTV#vASR`aAvOx_>Eu>~9AB<7I7g^-lgcyLu&moWIqPkFWHP?f2UKT`&KAanw$Si#q3p{&ID5'
    'rgM%u!^^Ah<d2KXm$k93V!PKrUhlQn;g#<1u+x#x-{~L6y>4`S$$#ws_HTdr*B}1tAK(1(ci;T$5C3Eg{U{hBOzk}XToLvAKm1W}'
    'MQ^s2;qUk#&;JDeC-Ogu|H=GM;eSkre+~XcCxc!F-3<B}bTsH`(AA)?MPG})7JV)HTJ*K(Yth%DuSH*rz7BmI`Z`%G`a1M==<Cqe'
    'p|3+<hrS+tJ^Fg|^|Dj+_2}!-*Q2jTUyr^4eFORi^bP161Wf1~&^MrOK;MAA5q%^2M)Zy78__okjL|ouZ$#gSz6pI3`X=;E=$p_t'
    'p>Gldp>IOpguWSlGx}!q&FGuaH=}Px-z?Zg-;BNmeGB>)^eyOH(6^v(LEnPDMMy(wLx`gj>ga?#I-!qF2&5AV>4Zc&p^;9Aq?1vI'
    '#Qs>>#}jEjDrx>v*zpropAP&l>%ikbGKvzJQkjyO(wP#PQks&Q(3%lqGgc{8MiF{5LU2YX&IrjFp*bT&XN2mEkev~_GfoL!MiJ68'
    'LVHGt&j|GyAwMJZXM_NaP@oYKG+xh*j3QKMgba<)p%FqfLWxF5(FiRXAx0zAXo4O;8AS-v2t^tpNh36AgeZ+rr4h0;LYGDe(?mVR'
    'GK$cq5#lsLokqyh5Pkl!Ds#j|dLHeJ+mHOu8Fv}^pEK??@;_(5?udWSfZ`GVo&nDz{yhV-NBnyRjF0&D3}_$m?-_7E;@>k6z=(g('
    'Km{ZIJp(C>^zR>=rb(+DtUg1#6tzcm&e_Z3t(UA~wwJ1ewrABv+q0^q?O8R``7<l1^Jmsm=g+LJ&YxLpy*1wE_V}#B&YxM8oj<ci'
    'd+Yh-qCe@k^Ji9b=g+L|-dgXg+s~}~&Y$1)w;ov&o<Flfe5?D#?f0x6-`dW2d%3JE&!1UmzCGUU?Xm`aYrFf~<Fh_}>v?y3eXL#I'
    '+W!7FdRW(P&rRhf;(7nOZ~n($|MtKBI%(ZMzMpB_`~Z#n$8QE=%QVFQd+~hpjQ$^VMaOvY({BA!e@CnJSi`#2xNg-ie|p~(uOGiT'
    '*j4*OUd^*^w709N2V$M8R%yL|v+9`F5HahaR{L5HzF#k4x9+!YG^}^C&+8Se`i*w|i(A2SJ@{_jYFzJTShwoem>BGOhj)PQE{Lu>'
    '8oOTVYS#yM2O#c1&${D!_``Z#V_rkhtx11PMt3Y>_0xX6guTB5rfGfr){Xpn!fw@XJ$k<;i1Ch&?vCNtkg@A!b?bg^jn?&qchJ^)'
    '@#7uHzoXR}%>BC6;g0>R0e=U^?qFJX9nL#g;d;||Q2ZU|b<M}_An!Fy$MskTe+M&l>z`I*4NE}0S3~c|Z+2z93sx}Kpj-o>td_Hj'
    '`2h@T7uFqFD_sq?hGxlk^fV8k#~o{`>)+f8E~>n;D6T4^{w*5dkGII`>d*iGfByL&|NP5e|J$?O5C8n(^z>8)HutpX!+j0%y>p~{'
    'TXgrexNnI5{vpi0JBjxVac=|T?p^2oW84SO{(fNGhxgv!w+esn9QSUu+`H5JC%Jc>_fK*!!0)At`zhysoZWv2?jPc(&+7Zh8YG=v'
    'S@)aGtDFA+fpMoNWcuau-ZL|-M!8)5GB!HE`}^>5QC1_~M@akCoX30n9Pi!R`Enu0`<A}HwbPBW9}>{hnCB~?U)-^o6-ZaVtQ5`n'
    'GgJSU!sxft8V`T{pFfZ>`Rw}7fBMhA|K*?m^&|(L<iL|0_?mJ6j=aZ`13w4X|Aq4SPvPkAC)eu}Mm`TCp9ucy>3#zL&jS7$Sa<^e'
    '&jS7$$NBLNzvF)8Y#7V^-tsjXFQ+!=Km7GyfB5I${{COj{^N(g|M2gB{^y_n{_GFKp?|fz{^PSh{qWbn{_UU7{^ifl{`~L%A&TMJ'
    'H~vYs%4*s7QOQO6<RyPRf!*Q#y!?#$?>CxOKz#rKfBJ%s53nUy0QlVFaPjFhcb<OIR#_jk<EPK{5R>l?zv`hpz+rm+smpplg>!#='
    'SUBQvSA@*0TbcV=%-)#$xz4a|^bkgLM<DMD^QLjV(EHidXw3tHEl*7C9;P;!?g25(!K|ZECOjanc|Xs*BRVy%XZzZmb-!N1irF64'
    'p!-PvyT1d5>p>6Z4oYADWWobt!-sXN&kzEi9AmpjdfXANepG>SM^yVkGUkp`@W2S4eTZ;+NPur|9$SPV1L<xN7QU3S`}Jka!+rHn'
    '(&MqE$FD$f{(N432jl()a@Mb|mp_a_K1fx6eAb6#%O45(?gRNnbNj-rwAK%5_*0{*;LitI{_+r}+TWFj)qQ~V&cbKwK-gM_4XagO'
    '<`1ObeV*58Ty1?FA05gA@})k43xBdhYq*0b>}t>JiSFff{T&8gUQOPu2x$c=#pIkK!|L?8kB<EICVVJ^Y(CX;^5Y5=59iFYZLY?>'
    'BQNcP-20OqA^SDtRzuzq%NW)Ky&7m(q2RD;bCDmA!?|A_)59IycG+f~&32wXlpXq0VfX#|I2`kbvPC~GT7Os*@(0~J^ES)5KCDvz'
    'L>|B-Y36%f{&JpI*#k8<|HsGm<$eio{#R8>zo-WKGi#Dh{?@xr#HVOpxW@7n&HA~bS-3THXEf^*9nzm2^SaXw^eA@yU5?_<vFq<z'
    '>0{gOch2!R_W4J#(7V?03pTj#aEo62G>Lr<TfM8pmkd??wEeqY&HL#>;1j3)xi~HEU;g+YKfvRlXl-PFf&u;M{On^edxs_b7|32*'
    'm7k0LwVn5AdC6-lW@&u9pDyV>o!OsFJK{UhPvHN#z>h2FcLt<BPNn?d&EYR>V}CkgKf#;$nQ2N7cGW$i>H9%H)K6+#9}%a$+gpQ5'
    '2EW`>eV0M-RSx<W$SeCi>G)Gx@cXF<#g7j(Rv%Zb`l3qIuY<GSVP$<rUhVf!=E=_}Yv1)x(Es_MkLoy2;QzV6k77hmUeNy=y`T?E'
    'XuY4je>zq5(cGuMq;%AW<(~d@qUkTj6}yJ(*Y(b#{@GW$R`lhV%VVcb{>qul=cYjZ%Azj!oa;OC7JrVT{?!r@KXlYl+VDq|tXy<H'
    'Tl328iZ5;Nm~^$tvvrov*@|Mf>v~gbui05yFW>>ZV!OVpaP}#;d(PIyw6^PCj9>4~J-V5Gw$8<R9PiDUy<`3BpVqa_?%+AwJ4ypS'
    'NXOZ(>G6JMbr#oHUM>9WjuLk3C2iNxJzG=S7t@-~?kG2S7j*4<cV9={+OCiE9qQKh4p#P)8dwAM8j}xk?`-b~-K`nT9oSwsDvxv('
    '=j^T<mv?}(tUIoedlw$U14P4_DIX^-YQJ=spA*$@@2tQ2per?xq{{t5eBGa>Rs9_9?@zzH@mbZbewy$<;7Uh9<sTpQy7I?f_qtP~'
    'dn6af$5ZpK!VLNK-P--X*cI~8*KqHG)i33$?cHnp*{{<+#^u=$8E&XC|CHwcl;;1G=Ks-Y{ujg3(Rn&LPe<nyj?Vt+=sX>rr=#-;'
    'M+YDId8keMrHhk%B`3^pyYR?emX7=F*XO&#a(t>O@^FWj`||bw@cpnKpZLmehOayo`TW0`V>~*0*j;_Rcm3n;2>$$B<-=ut9<7Sc'
    '*JC|j$=0~52++eE_`mWRo9j|PkE$=34^-cIs(1Ot9rWwcJP%XG9?k5-w_2Z=?Gv;8G-mr)ob@ZS(I=kyi{hETWCn?kS3hp9$FH^6'
    '@1y66{Pv9Fhvk_3s<`Gap&-!{cl|Ql^~>jnd{Pwe5-fM0eu1A2dL$3xAw88%|HwN-PkzUrUhgAc8Ty6SHUAtIv;P}G;)nnG2Y<Sp'
    '9{=TU|M}N{{Ri>}HiruuHvP_X>Ehkh$)@k=+59fCLG*gyzfFgyv&)N1ei!oZ${+P}@qaMMpsRi){ST(dOZ`y#AFQv>@(tZTnEbo&'
    '$LR~oe{lA9CjN2umEBnje9Nx2sGoNHgH7^7{n%jpp??0Zb*Ud@Yc1-hnEv2g`vD)en72T=_xTto;5%=DvZV7LP$+Z$1Ip>me?Wn^'
    'TR*&9j=|-m=0Bjo(fkLLrkVeM0w=c?_Wfd7F5EDGMk#^$54hZZYlw?Q3tULMwYPDxawLb(e?Tha?Lim2N7~A*-R~E(AY^vyhuvbO'
    'aOB(q$9{oph}HJPbO`=noN>Rid@}y~U;aS{=`Vl!;ooM*;nrDb2Lazr#@CbeKSusz<v&jT<K;g={uAXtN&b`NKjj<wP&EFPv1AMx'
    'LPn4Qq`!2R-sp_x7F{jIvGNNBuo%E%0E+=E2Cx{wVgNLB7{FlwhXI`YjsYA7a2UX00EYn_20#;!0Xzoq7{FryFK;n`#{eD!cnsh%'
    '0A2<R5HLW%009F83=pIZx(0L&=!y{|x<+)3=o-;AqH9FgC@nES!~hWk;D>|(5(Y>ZAYp)n0TKpC7$8Y|43IDY{>&I4V}OhSG6u*P'
    'AfszW*DUSPRsKQOg02N!3%V9`EttDtfPw)E1}NOO2gUS4guM`9FGSc25%xlay%1q9MA!=v_Ckcc5Mghmw+<;|$PhAu3?TibyYxn9'
    'GzSs(Aj00tFBkws*n<dr5Md7@>_LP*h_FXP5Md7@>_LRRlix7_h_D9{_8`I@MA(A}do%$N_8`I@MA(A}doOP>0En;$5%wU$9z@vV'
    'C5W&G5%wU$9z@uK2>T#y&=o}3g9v+!2qNr3gguC`2NCul!X8A}M`?)xK!iPru*VM|!X8A}g9v*NVGkngL4-Ysuusw+1Aqv7{0SoL'
    'L4-Ysum=(LAi^F**n<fBEbY-%{y|p|VGkngL4-Ysum=(LAi^F**n<dr5a9rd8H5N2A;Lk3a1bIKga`*A!a;~|5F#9e2nQj;!ANf%'
    'QpS)WWCR&N`b&4|jm~HeA{;=3gOy(}0Eln^5e^{20Yo@}2nP`1fQBH#0Yo@}2nQ#>V*n8003sYfgae3h01*yo0wNqhgae3h01*ye'
    '-eLd{;Q%5WK!gK`aKK9t;Q%5WK!gK`Z~zewLE4}zh;RTA4j2(cIDiNT5a9qK96*Evh;WF~5(9t;2N2<aA3%fyh;RTA4j{q-L^yy5'
    '2N2<qq&)@z5f1nhL^yy52N2-^A{;=31Bh?{5e`|}qpSRbt{}nzL^yy52N2-^A{;=31Bh?{5e^{25fn2D5spHHqY&XJL^ui&jzWZ^'
    '5aB39I0_MtLWHA{-a4d=Aw$RrGJy1#?$R5b(HulLf(S<|zhD3m;RqreL4+fSa0C&KAi@z1L4+fSa0C&KPJYJ#Ai@zuID!aA5a9?S'
    '9MJ?sID!aA5a9?S9KF2703gB<L^y&7M-btNmmtCsL^y&7M-br%A{>LXL01sr2qGLYB8YGV5so0j5kxqG2uBd%7^Ni!01=KL!Vy1!'
    '2uBd%2qGLogd>P>1QCuP!ZAsE3;-e=@h6CI1QCuP!VyF`f(S<t;Rqrev$RK7`3GG=gd>P>1QCuP!VyF`f(S<t;Rqrei3n$AQPU31'
    'N2+q><8-sn*2*5Bk9|lJi%WehHf?IInwXE&Bh1Gs2A{1GJU|)vkQxw|%0O(Qd@dKwM+!0XaZ>5C<<SEq(T8METuP$Y<j9;Un~&t6'
    '=Hpo6XEVbGu)z;8z_?_Cu`w|oV-E=I0f9Xrum=S8fWRIQ*aHH4Kwu9D>;Zv2Ag~7n_UH-->;Zv2Ah37xI|c&;_JF`15ZD6(d&~z2'
    '>;Zv2Ag~7n_FmqiD<H531onWy9^(Q6dq7|h2<!oYeUP{43JB~0fjxe}Gmmiq1P*|}0T4I<0tZ0g00<lafde3L00a(zzyT0A00IYe'
    '1q2R&zyT0AIQboe0Rjg=-~b350D%MM0|X9$zyT0A00IXuZ_yPHH~<0%K;VFJ0f7S`Z~z1jfWRThTXY2k4uHS`Kj0ZsLA+yM(qmAc'
    'WGt08DsWWdsK`;7qe4fdj&`%^eUP!N{wQOK8^MSl!H6Tlh$q2_pwdZ2R2ip&PhA=MP6iRFWkhTlky}Otml4UOD;L>iM0gpIUdF31'
    ')h05Q$T1^=%!nk@eTx(`BF2o!F%wi4YZK`$vdoBELBS=C1tXpXBd!G_UIi0Xlxs`rEuIC1vUnAYxD_7avhlxuH%<IpH%<D=u9@(0'
    '4w|`ZX6}^1Nw1!7RNj)xR#J>fYVt_=DX9x1rEa8_k`&64LNh2@g`zj8bMw1}D*2JV!PKCzmPYpDWD<VJce8c-am*b*<cUAl0Kw-i'
    '5a+3iPV@!y+T&C`=JGL7l~Cx5i4K_Pj?HN-V;<>PD2E^G=lGC|BQz^~t$uy18T%m%HZ&vavq|cM7z<wuV_}IgL=6#NLxdS2NyK^)'
    'Hb+Dial%_ss(FO!3t`rOHT?RohGqYCaqT}F<1Om=gyi2Yl3OJC3CX`*B)9B<CnW!Nk=!yqJ`>45h2GCY>QABb^N@Hi`aTb7_n>Q@'
    'U2`YE*Fd%(0N@&E_5%P}1Hn$4U2kx`DLY@hWCA@J%{&%`JQ_Wu$HaaGnq*5YfOa>=Zv*6RT;JZT|FB(pdn5isR_8o}>xuFI${7DV'
    '+wKW@ziFv08I+%l-VaIYPeb8{r1Pht=_e)fr=j8=DJ<EiPbcO#ePSe|_%o6GWfA*jk@*!6`1nYB`6SDJ`?3s~{7*>!O(R)uXMA}i'
    'e--?GJ=A_ZtXA_uZq0l_f6!x{!>@4yzuMXRLZ?n{JUz;(`x;F3Yp~O=#z-IEQ~Ujsnev;LlemHSgyY{dj^$SA6Ow<^NS2$zKY`xQ'
    '_sV^~C+?BFZy)V(lUvhYDBSoZ0*zlR#P|h+i*ghFqbTs9?BwAT`B1j;a0>k>_V93u{a7{t4uHddY<}c%_XSL+FJ&%$DHG|7nnj<^'
    'kEQCu$DD>wq4!Ur=uf8Eb?i?nNjzy5zs|tz#qfl@UoG<XPssb#B2TL6Jf7fwe3AY5!ukt|={0brGSQP(^p&)tdG)6U0{y2npN9bV'
    'r?jAl0P#^SQawOBk_uR#M~RO^caKA5U+%tHs}A-EVESADJql<Z2`Ew_?a_l{O|n0BM63z&#}0#qFA(tx-`uV`$Mxp^@y{Rr`IkR&'
    'tvM8x`|+EN`R4cEZ2$JJfBA2Jdv^A3fBW0N|I45M{@<wk8UOy5fBpad@{hm#l^!>i9<Oh%&!4@zc-DV=b9nuZ8}YA)zU$83y?JqU'
    '#68%Jqu<`XIzGF(o;BnjZ?3a(vAlB~?#FIC{r2Y7O@F(kM@t{QvxCdHL%Rt)<o^10!+?e<Q`1dIxO=>byvpOd+YMuDXt_ER=1y4T'
    'aJyleEi!hvMZC$p$k%UfH_T{QtSwzm9=Dk{1yjGhz73#K0p#9o#}~L7z2PFD-}K)tAcSABahqd`xDdTDjD_LtVbRn~O_RymlyPx>'
    'xp}>4Yjv$Amw|vPD*%pmp?@~Mo44hkU)^*xqA5Nl8J7`qcyYBD*3XuZJ6zV@@aMv5W6RZshEYFe*Gbcerpv{~S<{q2#~Z_F*%D}P'
    'Ll!>+wBhXhs>4O<jbY4euJ_B;3c{>&T`)P}F7?JRZr*%%xoBD{Zadmqz|HIpS3~{g{PoTG?Y6d?(`vz;@(tJ0rkyQl(3&}R6Ry{9'
    'xB&DwFV7YSz}gvMr-BRU8?L+iP2VpU;AU8`PO!iwUugl~FPhE}7JY{T4;u@FZTE|&Vb+w|3J4fA=iB`Pxwc`<ma~|QtbkE-efWO4'
    'u`~x)5<;*Ut6eT3&z6wuu#6gicJ=*js>RPcB7|<!ee17qyME*3jExH}G(Tgxi_+A97hEk?>&%QB<cJa)8wZBB<QvY;(BO(}%n2V{'
    'a1J+TgtjX(DvsEjWz+c)vz;zLaC5$F8|G*zC{(a<a>iaPn$FKywkZ9uag3TbXN#6;&gOy<?k@PcSu}0<a$$51wH`K(5z~LaY+B~j'
    ')d{+Mhr@inz?bi4YbDl=VLc;e+`U~iH8VbEI+V!Rc%bWIw^_8cGiZf^stg-1JaO~(Hg$RBi8gs_j91A8(RI0C9p2C2<sxcjY&>Mw'
    'tCx$WVNP5slf2;TyG7GDr&()dRFPi`CgbM<%y&sfWz>9ovw5{>o9Bno2}|^ho6CMd-oDJmO$RMy6FRJ5yO7D<d|v?-MK-}ehAucq'
    'FmvjPauNb#X0v&_pxH1xVA*8hp22};-n_ak7xZV63*TOtm!%U0f>A@b?HAaU-<fxWLO>7#3yhq@{@u-D;7|wl$9L&nlD#o@HseB$'
    'g*ZD7Hi=*eg0lsi33Hr<F3Q$2nl8>`$g`%V)14JGT>`hvnx+mFST?c4TF3GBzGByOlpz~0+$%zgoiA1pXHqQNENnT75=d=t58HG@'
    'dvSdE232Y{k&*M=;ewX1owsc`dw1^4IjYHQqQ}k_>?8U)g=Jykkx}#N+ufpRm@`l>Cnqv$cEf_eq9ClKWYTwZc14+lhVk7(=EgJ|'
    'HHIJyVAPzQZ5NCt)<YJ(`PN#|C?ha-xlsRNX;MF5g&=g07&UKh7F;oP^ONJWDG+wpEqGEgGge2s7a*G$HT!qVVdsY|NX?{80%o>v'
    '76hK$jKEDNpiMy7*@6_4pQ$rBiQ<tMHJdj#3o(`iW1S6MBIu&(&L%N#-s~3VCdEllBWW+>m7sVWuNNpwX$IRI)B<4K46ok4yiK<F'
    '8Kk)X{_3h)3y@?t!^L9U8DlIaqpsE_Gj^^o-!9}+o~a%wSb=V4?7aN0UmUY6NDHK;#pF~?$rdEmFU~`@GZHfe2ID3JYt;YQ#b(~N'
    '9kcV-Z_PU|<3e4$-aQ+R^QQdka79eUgwPpbb~!*3zgjTH+!$kc1N{eW!5H7{mX=eTW4MFWN>l+o_`4+->Pa&$_tA<fo<;xoeqjve'
    'JV%8yvb6$`y;vNzT;`|Jh?!FuKQBm5-fmmES=$r^yafVnSm><9%qY@#C`Gj?jGi}3`B3bvY3i&rbx<l6bW_}{Y0jvbwJFd)&K62V'
    'nd={+2)8Lg$o%!fF0DB0w&$JQixyU3KA$i573UUYl9N^#H?LkV4q6e6k*a9Ym8~FP)NGf?EqTTkTtKO-O<~M@J6<hRpi(C8{EqBW'
    'wW1j{-+ue;`O=Q=CbAi$M?uvM*S=#oy#8+4TCnDVPwp7AFjFp($F0wnc4TK~t+NX{wi8AB&9e23OWwGoMsnwcL43Px?dR*3x^Ww>'
    '%SSZ)>h;C#qb$FR)nmUs`kk%>r|xP=Z}K~>JC)G&B!|SZ?-F+Ta<$7e+hvksfEz9bNF+0^Hks#)&7ex#hD!#LEV$k+8<)8-h*k`7'
    'xIWG{D=&v2vMWT-+ms1Y?_0gwk8@m|bFXLlnX%8i^c-#lxUSCMcIV?|$N%T5>q^?1?wK*qZTB3y2;5gw<1<^`S8d?umfM^B91a#Z'
    'v8MJfLUm$E`}E8`_s_cU+?MB1IpNEi+8=H8WzF03`1D^}U0PH7t6W`Li}uz%N0mN!wx;$swtBXteUSCI=e(p?9b8lUcfLBf+I|fK'
    '66N~f;%aTBxTLtaj{Vimc=jxu(%AVq!Cuy)?~c@#xK9{;b98O@2U3siX$a3cT6}sA6RNtqOeA&B?JiIaP~2U|{+u-W_$)ooJVobq'
    '`RepaCI<iQ_AE|Lui66}F3+veO&B-tvGumOzG@FfF@cCsVAIS4rah!Mt9Zb;y`%OtgMV(~bLbG&4OZyDq_d5bQ584X>t|MW5T2Xx'
    '9Nxt02`h%Q?C$C;n&t`P)>w;0>(G2*t@X>;HD4IFuEZ1SUh};5_oCv!@3mM-nLEVL`da+e;t+fN%*aYIw-(5&ON?97|2ATR!ucgD'
    'sbj;0t6ni_9bbzt0{$=AWnDMi<JB=Ht<Qs~MR0-fN&=5<XDe~D)<+*F-&kuMX083K^_3I%0UTl68tbOB+Ev~WX02)2u6f8vByJCt'
    'A_@BhA3yVxjlreVZO%ktwyE_<IDV$CCe2%KyU;vkXNQZko2zGDmLwCTnaAvCoGo#4F79wE3ecMkX~coxnfY8XlXA^*#$6&wgw46Q'
    'LEs|W_OCWq8=6Ilm^WT#4z!~+CA`@jyVx9Pwacl+*)2-->P6EIS3<||mr;kFSP!?4&#hJ3R9D)O)SWJGHsbt*(`?&cT;J>nv=z}-'
    'iC<l6+$Fr&yb^CXoM+laft8+{q`*?wns)hC&LVth+spHwtkYQGYiaXg=3zV9LI%WUQ%du&ZFl{vI)`NzD(6Q1Y)5CjyEr>5uKC%}'
    '!Bv7PGgq57?sY-AH9k!YRcao(nY&G!58|PPV@;bIa8Tnb7RtQ+Ug%YPZo~Dl2cn6och5~x2&w1o7^02CW+M~`AKT^C*<nM}l;lhh'
    'Oreuj2i!3v(mpp*oLd}l+w<d=NW9T3gEh_kaL16z7}&g0pWVc!y-Z@6P21eOac{3pAX$*fKupaWS9^f5Q;(>S#Y^gut@qX^p;#Yo'
    '33^~n&Z)mGiAB|f3r_akBnLrCM6vqkbl>v%6Y<YgVGJG_pVI-iolVcizx~gDHgDZD)h;)0T@#!xsN^K&znePdTLN@%o9piNU;gwz'
    '{`~i6X6Cg6kJ|3uiSh~moVKzvhn;fw@5LU2Yff9a=Bt}zyy$pv%4sV<^U~>9H_kUNowf?iKUZ75mUyZ7=d@L9j=8$};Y=(fc;K{E'
    'YA(6+x98WY#F^%j(^h%rjgzrnd@ptp{BGK+G)LUD$KE$bT!m`aoN!aKSA~^Ql-vHuSnA>WaIxJyd-Zk^PGw(n^qy4Z(sa`O_2%u1'
    'x7*u|)iLcm?@RQl>7<97>ucs=w;FqC9HQ@1v8YKb3hCj+@b<2_Y9M3Faa&(?(aYE4tL<=edtAu88xzTj&e@5r#eh5>&JP!KUJkoO'
    'd=-9zoL+7R^(oMASC?nQ;jq6wa+VcGYsvPuwyB_pi~f9jcDUWRXk%ks3C^abgYLe&-n=|)pN&_`P%CUFdTpiHvS^_F^^4v1W-&P4'
    'a9Fe>lZ5mX*YwZp?fy6{k!Tc)WOR{SEqPKtUz}aP*c=yVtU*`8x9>>c`=)umIlMl;y)*0zZvg#IBebG)UhT%sa^*qCPS!{3{G@YU'
    '5y|%BcZ=px>5ti~4}u`*mM@N*>%(S&d%R%^^d<*uMI=fD_N?#E4|dT!R~n2y8&gtKJCBF$=DJ_5U2$p%=yY<$Ck*xSYTWkwB?aIO'
    'uieHHLW;I2o@bl>{Oo2yig*KBeEI$O@?0F`P%wvKJFK>91*ar;x)?esD=(_&)qz!p!(#O2Xd_#GkP7vpd+z$1^Udyh(cD*OOb)?i'
    'F(sg5o{#-u*!GL&VamGrHk=99llbiA#jDK{kXFwqouSSZCtL*@=9|NKczb7zR(wbtBW;@J=GFH^Fs~P5D^80gQRzB04&<1=*}vqY'
    'aj^{5j%Q?TQ`_t$LI$03zujM49v02<hMS-HcW>Xkmc>A$q|y7Ee%aj2M^|a!^qlVd<vDfroO)szF^S6%%H;WGcrhFnB;Z^HBb|?='
    'IE5m)8~QhI?~L7A+G6KIPE+jmd?0Gvc8js~oMsaX@c|(KQ6SH@1Qxewt{5%Bm`p48O|`tZ?r+Ac!vZd;!lh$At*R^J*!4IZF1}xk'
    'UO8{^)|D<!vE1?cMSncJSTrvR7n3+TDAE<`<8ZZs47e?b>N$0Wcv$MDknVww_4V@Ld7}r<Gk4US0qBk6=A1yYI2yhMokI6w=fV`('
    'B~HIP&`G)_2;@x_1mo)|xO;W|s=sjyK;jJpk~3t2e49aZ$=&sSyJ8w%u*yMW;_IDCXu<0B_U5>ROso>G(028lLP<Q1`*$QE7hA2;'
    'w6mU0pDRs$d~>~B9&10(SjtHafFfYLezD|^eo9-~z8<ecZa`H$o*l0H&0;B`Mi4ufOR#B*`d)5cj@!f8?Tv;iPYJ2k=UL?U)%E##'
    '|LuZ*1wE(4x=7J6X$t;ceY<~m*e{?S^qd+NV@^HxP$!3BbA7!$*1=cdo-*0a;#7cMxxYT^U%l-Y&BILVD)JQ&(Kv6;M$#M>I1PeR'
    'GBZm~A&40TO>?~Nx7SN~9aN$^<A`fm@pM7sq~>d_CrmmL9c_&hA@n<twX$7mR6!+bCbm?9?|cgizaeRTbG=)rc3tJ|k*HBqmJk8n'
    '9q1$-7J5w7lgbRnC5_P|2z>sE&eCR~4aBPJrc7KwLj{NfpN*UG)v9?-O|yoe?u4Nt7X0G)qCaewi&3mvK6_EVStC(mVH-EM2{=v>'
    'ELy(P=J2Jz+K_Z#E=EtPE76fLB*qEi!skHKf}uvmYV4fvvddG%_;4|fw3<cptOQIfCM7SfNQ8&4$Nu_qDN|#WsiX*);w29bAz~W+'
    'X1hdFQmh7|MmnjI4~Qu7f%wL@Ul4GrSWTfsBIB;bj~S4MeyQQx8RHnf6LS*5Vy5?gc|1~0u5mtQ5(mD8kx62cz+SFhv1&5O?3$p7'
    'nDK7dhQn?_?@9S@qLn+BZP%j7#AnueQmXvdXZDkbB9728n;fJnEt)5ldnquZ5tR_p<JbN9X2oYxt~5jHBOllnR~|S0&2Y2ek*R6F'
    'u5e~bYN6#9XZ?A9vw%$QlxAo-g~V!U(Ph%AF8U?C=cfI->-`{~a+3h__2y##ZnxNK*0bt}+(}|ftPa6sGV@5CS%`xcW+wZQsAd*s'
    'zS_ULTruoiZBxdR?Zgd<nDTLR^@3#CqIsz173ahyVH^l6zc`-1T5*CL=bRvjqL0Y(6&<C1iKeVr4P?%Q&e<vE%oKUP)FE?jwjR-!'
    '75g`f{>v1$Y_3=(xmUcaiNi=Py?%Y@myDxy74DtSA-O5|yc@1xzr8d3EGlRA;s7pIm$Fh$O0{T=EsnDnS_u03jLfT-O9Q)FsWEv{'
    '^}-Z}=FzvyjTS%OXs2OwNgwF9WR5MG>q#YzmP3>ho<!VvcQ#yIEO}UoRVidk(n^|0QJj+gupGVFsW5~kLWTG<S+)JEC4m*?zm64%'
    'qFKENL2t(Evx8k|1V#1VO!kcPLCgpVLeuHEK3p$Yq`XqaEzpL`6raAJ<3y_7!VvCEHB^X|k$sS+5cSzrf4*Wh@`^bbDV;<O$xKmd'
    'GMskrme`V4FwH}z8LIyTnIWVzZC@>I^G@(fWN(UNyK;(Ezc_ov7S6ImJGChX>s^?_)fbnqHfQTSRBqg*l5LCP!^qiwcX6`@Yx6F>'
    'dv`JV9HyZ4NP5}zdefC7v!{dfN;?IvH%t{vM9V8++WKg$M!Jn+dV9E7F3Ggay7*@AcQxA0=-eD{_6NJ%w0fGDmU1LTL1#cx<l7rk'
    'c2^ym78N@aau!;GOKBMHIQ2%oO>kuRl0_jka4Vf3j!T|m)ZZk|!0f#4K-J@$>-}N5<;gek?#2E}66p*aOT*Cb#={b*>T@zK(CBoO'
    'b;}7@AknghQJSzocDop12F|73FuYmqAkQuFqFEJ2O`rYt`<vwsn(g!M_~I()g$b+1&x?LsZrOHC+>u({ZetjDs^~Nyc1!=Dsj5I`'
    '^GB2<0~gcgX2(629jXOX=ycks9K-5)7Sg_TH3F}iC&XI8GTcmL8T8|JHG^-afoFfKpbd_u&DnNyuqz*D<=DjBElF-&5>)2xd4Fw|'
    'cvde|7ufAjj9AbHchl>`s}<VR55;+#{<~Oj!@%8i99~`BT(5RePcz9r62*0V5yhkeU9Wbar<oL*SW@NNPRyaJ4H@VQJiB(9cVuc^'
    'MD2jk$YFP}+<`_u)sG%Utub&q5%<DbUW`BK&R5@MU(|2xs$5v=O-4OZU37G8q78yHxSlq2GH+H}b`{IHD<pcv_hR6D+FZUPhPQG}'
    'hnZ|ksnaC^m7D#$o27$aUSSWz#dq(n^)$l)MMr0Iy?L>mA<g|eDr^m0P|z)3ZI?5&{7hS3c1I;VIH8Eau;RU(L3OU=ZEzupqGQBC'
    'H4q&w6-HA%RHPJnQg&1>z!60{;C6qx>R_shNpg+kxojG)sG%p~Bf400u(Q)vtf*?ni}5=gH)qT9?CW`EE8kho9r5CPyj~q1KleW7'
    '>S>2N>hkQxb|k5AE7?<CA=}fL%$A{az-eZS_P7M9`l?9ilI7!!vU#|qHg7lA#Ke{zQZ?N&wN^)nfm>?383;&kmK}064I}B{n!#@1'
    'n0j-0F_2}j>`<CRp<@S{bbxD$w8H(_(i<YLzyL@?GEu`ia84bs4#XqPvP0FwNGHiGaHWG4(}Ot15^Z<|)7-i+UezTzIiSR3mXaY<'
    'F9*#)TTq@PmlTFIlMJ`;Y?F>3)Wf*nt-?W5JgBgKE2Ir46`eMcJj?SO;-ss_PVW^N36*YMZ;r25GgMuhEDA$%M0X&mgReH%=c^8i'
    '=R}Rka8+R#gl~`6o0X^v)!*y_tA#56;jFqIUmP!1Scu)+L2{<V9mNDb-yOC~zN(%qagngll@5^4uU-+$ScT2hmLgFbgYAqi384hG'
    'j#hh!ZWfXuMMj_mrxocKB(2^p@DOW+=&XD3F1`x7q##S$ezyu)M&;b4kOU3t@HKE;5sT?>RxAdul*O>?QD;ap!{+>Y1*LkTN$uf#'
    '5DHe9%E!ay{@wS>4oyObLI+eFCMK%8vcFnddGboxBr!cTW8lCVulZKja?{nfMJy?k@~xa(u0WiPhmBjZfa=}im`lfVa4=uLsQaVd'
    'CjMVP{KvDKLB2GvkNk-*jdO4J*WYco&$j2si)Y*a{QF=3^k4EhdU<nPzCpMfeTHS9`sTQNgQz$9m}Q^(=D2(#!Ebc3%TDdp(XM)#'
    '+RLqa)mKNi>Sb%Mu<BJ`9mA@ZtG(i?SABJit6skLN~>P=)iJGlh1x5xdev9Qyy_Keud?b@UmeS;S0fA4t-3YBSiO_Lx%69uu@Q$^'
    '!?BTvUC-ADWby9(^a<o2zZn2-d3-;8F<4*d<NN7@v(pQHd_R4pRA1=h`{~1$`a&PyPai|n7y9^q`hcFk(8u@F$K9qE{P=$QaN6{O'
    'AKy<O2b*5-<NN7@T+<7Fd_R5cX*T($t_mLr(iimiv|N1bX6nI@@23x>OfUHH{qzxz*^8HL4*qb%^nxGXPsRSH7yS5sDlR|0;K%n<'
    'q4F6>Uey;UdOh3V_jS4N`uXo%QuqB-T=%XQD9Ad!;K%n<G1A!-7qtfpXihKq@%>cjaR$S&LIgzyrx*PAek#m2z2L|9Q_;8C<ogPB'
    '6coGb1&US8puVrGLZPME247WrLouM)i(TD03gXPr@vgo=@s}Ao_LUG(0A;o+dVEiB{`AA2{`;T##rfvIh~z`IyE@tQy*y~~j*<P('
    'raL8@-O_Pq>no?zSOUJ`9{8O-wel(V-#dHi3jNIr6e$#^F-kR*-^X;?Ib-OFjVm8dxugU=u{LVlX>!5SjI8pfcxq{go|U%H(Hs8C'
    'yX9$k2<pO$O*kcOpH?ukU6lN3>MTEDaru;#)}-P@7u2&r3hA^Y{Ebz+)6Vm7^v0iB!%H3bg)66~NQIZ*J9<NV8#(4g3nLZ8@G@z4'
    'E@IasD{`ef&7_#|qKWxNcglYc?B>tmGy`fRH<IhHfdS6e>`pVUpC-(PRzkDy9M)w{EiXkLIxKl)vw&2Z)8u&pJ6*Fij0AT0Cv)l@'
    'O3~08OHVdf26mKbbLK%f2x7}=4w;*>&B_?cCW&TXElpN<GHypmA!VEa#7=)Qx=%~sMfMy9BJ1zeI)+4txef1^{qc^v70a)7b;&;S'
    'f(1&6X@f&IhJcrNG*$^2vmx+<U^eo|?3^(vonnuygcznKP#vg*mWN#^$hg!4z2aaW1I^+cKV*9vY-Z?Ut=(xz{I`$|<7TwjJ5hAP'
    'j;IqE*{2=Io6#0d3vYov{6x2kL4KhA4y~w*52ubFh)xK?phl<|F`n{1cdP;g8i}W%pNlz7+@{yblh7z_(_44yd6(=8@&@zmI4DU5'
    'GzhjATFR%6Nil(}JA>~0BAq(?9c8nQClSmPK#VoaWNB3l%7Hy*J`q_lDHcH(fh%EKN=ku|%f^9sI2b#DD`7pUrbjDq%mFyMNeVKb'
    'i(W966aMB7iADy!vIb4PaeM__Nlqn!q7Wr?L24oKu(?Er47;ryO$g{=_CRMWpck_SI<W~b2v+bfz-3h^m5d`iL}gV4$k|2Z0U$39'
    'oYL`je1zF?5@gq0wTbyj$cbT~Cx-dS$k8N%6_Gl#3U(v}k8|5|%r?*|!1C%Ww!_XV7vV#u95UI=GjTwO4oH{L5Tg@@`P2&UqvO}P'
    '2zWD*bVbQZpdT2qI&sICV8<Y#i#hhL;WdcW5M-k}<qcrnGfEX?5zu&*ay)7Bj78UXAvzER?}1ql#sL1PsYfxa4$PAaTJTi-DxYQm'
    'J~3H!r-qTjE*oPK&MP!Y+5pxGE<>ESW76PIvNd2%gK|3qQmRzK5EO<HiO3wE5_Y}vG3Em>V9v<o#1G1gz@!4sA}?~Nd*!RsQu2#d'
    ';7{e8$V3oQdN7ERnBr7AMldjyIH|+Y7Wu)%z2i_j9a%6B6@Vafo!@y#HjX)Z!c;7bxIShfW&$>dO!yb`+7YrPo_ffQK|miMc6P$O'
    'Ei<rQIerOzz;~77GvrjdqC30(nXUnNMO6|Okigkgv>3><t{5g09XdJ(q5+3VB$fFPG8CdYjZD766KIQp(X!)w2}9@A;doW`A=<!4'
    'M_Ha{An~7_@=L}ZOH!seq2|O~iksmTGgChPLuKV2$^hPxn6+3x2f86On2YT6ltb)Rum^BK&&rDB_2K|qs5?M48=8igS^%#zq7^Yi'
    '<j4jKcSsiIP94}suZm9q95R?Qog<u2T^r+~hcF_#!--R5L>qx~NTZICIjD4WIG4hXgG+Ms6g3Pip)Q`Y$GSv>jCH*!Gk8Idd~g6L'
    'VHUxHysHR}j4nNr391@%bUy3`z{o;c9#E3yU@@wLLAR=8pDh>>Kx-hlc6z20|E=to*F3Ptl;{Fv-s{SeP^Kt7Eayky*^t%+om5nf'
    'bPDnpg)&A^A}^VC=AqXJv&pA?J(7D(XMa?k#CgWf$hj!2RF)1Ss!Zub$rZ*1T2zvPtkKaJ90N8QY2P4f)li^;MHwoXf`bR?#{*f}'
    'bY`wLLU>0KE)j~FNYcbg$f)F)1GWcR381?X{EoK@k`jf1)O4^Y3oV8rrKdJzh&O1f1Vvkfi-;a$pa{z=8)QME2c^V>FmN_Pl(W<0'
    'Zz~#mCBqCZ$z)JK5@VJ4P#JOBXm{Y>dR`gV%GjZ*)w9a07a>EZh^MHEujp5RQ^p){dn5u4pOi}LfCSHNfuliGhYsq_z!H}h2undY'
    'fE58%!EbPE^)wh!H$)dE1mLJA;boW}LV*@jm~)G8XGlPnslpq;H<bjiIF|EK7{?&;OQ1Ih{yek7omKriiDD;6g7&21#%CNh;T3$O'
    'Ivg-1VgP|@kkortP{uZ%wLFy`#E8A?xdOB*R}+PoDKrZ!tD0$&MR)~C0W+-}OSIb#1e$q0(BbHJUaDxow(ud}abSpaerZ+h0Rs^@'
    's!kIPz^beTas_gVkQAUgROK9so}7^&oKe>eq1g!s3k8ZKbl)smtO5(hm8eMGO37^;umb)eLXH@<&p5r%)=!}kLnU|UKL|jgoDbo7'
    'K}cL6PpVafDi(wtdB%cJj1##l-~n)SGrZ*<n@qB@x-rt=AXY1!_#h>8Vf|`Upa)pDJJnT~sZ{wV6xswatR6)uj<lB*lLDNjW+_j@'
    'dCcmd%(9*VGR%<~z$;K?7E&0751IR&gOU&SNj!#If$52)=)npLrj!moOQlp;tlq;k)ALp!`IVe_msMtPwgwZ8T1#Z-g7`%UK~Z;D'
    '7_*X5&>SFJ#%ddu0$$#UCg^b{s%kF9-dQ+f$Eu^l^t~l?pewTS4h|<2B>KPbr~)1_R(&}I3|ILAj<P8EhBU9h3dT@Pa1XVRAsrxG'
    '2E{1YuZ&<|3=RspBz_PEI+-L14w&XI=wumcy+Zo$s?&pMQVu4u2I&5k#)|{}=%Mv6_Q$FkbqO98(eX*1f|-l_GMFr>sSJfc%`36H'
    '#e@Q3seuoUugudM$kGC6@&?tG3Q-F1igu1C<ps9D5~+%=?ZjI3IOyU?P!CO!%DiMMvY{phofxc;$Gz+|2nDGj4-Qc6U3gm5tg#74'
    'SmmPV3<?}<0Y}J-a8Y18@jU2`z_7uG9kW;-yDH5As@DkcP&d!AYggG26i$UNmY|v>jI(NUfS8;LREB`EMCgGcX=$jGp!cW>01a3D'
    '1dI{cbpW#)n;G3fy{`wB_alY~^o|}H9&v84QBfZXGlYTVE3xrVB5Y?7+&Fbg^02GCIC4bURy)rd`St<009M>XZ6G4-JfJIUE*-QR'
    'r2|-<tkKb=;7Do>7V#zu5&D9`o8ssAai@pUgXVVXDFlUAZ5+HqSV6XUF;m%9&I*qjl>fqvRxKB9iMj$tS~ahs7po2hDxR5myUP3E'
    'r4dRaHtgqBY`DTRaaZgD{#zO|!6IY@g3L(6jQ7AJ(Kyg1wFmIUv_nD<0w7>Rj&&5SF#UUWp(S<PvYEup;Gi#{{W91K={`K_f!9so'
    'W38HNRr747I&n!3ifRxsTPW(l&1@+L_23rP|FigM*;T}+gR7kUp0y+Du?JX5DTw?#{>~+;$CgIoATQ1r(>4*5>(mEJ6bPDUWnCQ='
    'L;V{|jNB%%IeOqZJZ;se6E(<cN8KGwci^r>y6hB#i<F1K?l{+h`Z6r{3RgzN9`GaLbmHqI%g~5!Byfd?Y}8R<jOnIBP{$^l?74Y1'
    'qqv{o+6<>z-I|3Z56`ApOirDY)V>n|6@V>x7M;*lm%O?u`Is83p2@-{suKq!dSu406`bPl1c?q;cEpZ6pv0J9=p}U|^1|U?Wzf*5'
    'TvZG>f@iVlN$LeCaFZMqa8Py_$wiq7T!V$3a)CF=d!Zgl-GX%5Qs+!_3`X|l%A*&4kXS9;^9f7>e)#`L2;;;us(GiN4Q`YP1k8{3'
    'N<UGmv2zFO)i=lMhwnglhcnLSon9CmL8VL*+#&K?F3|pa=5ATMbu^`cb5(#C;I>kvS?0$Fp`DHe?;gk2!si8TiEYiU`qo$zw^_E#'
    'QiD$r4}IXK@ia&kcv^pi9;_B^W_@B0v3HsfB;oVTdx^5*y`mjwU)$NF-Y`CN9$=L@x}Y8~(jAN_Z-OUSmW`6VI<oM<nMXz-S8>3w'
    '*7L@4=OD*37xipWn-0pPl4?RgM<3t}gHwr-z&M9fi4_4MJbl%f>QX92R6)3uXm<>egRI*z+c_xOA?+(MLA&yRr^Uj&tR2EA3TZf{'
    '^<EJgSU__ze;%!<pNN+to@oOYl3{yy9~kLy4DK11iG2$6xAH0QMGO;xGClqo<LjcpRTUM1)y_k~ku_0$Lh&?0yn%KYkJe;}NtwG;'
    'HxQqsO1;Y9m|yKH3t(nf8DI4fCDx6+d|3ce$Eky;WMQp)JqJu92i+l5UWh3+dX<&yX|f&?CgT8zZ`Y~EEy=Qb5gHM^RLK!2RzinQ'
    'dyh36$Y13eBpJ-);vrdv4JoX$dU$x*5ax9ca%zkmryy<|)%<lg&ZrB9e=h2CArW6VOmk5DLXZjH57n(guuCe)4k`z5_Cy$^;?=Xl'
    '3>$b=f`Z!vtd9)~IZ}i2V6c*1;~56bM1&hywVrwDH0r(}i9Cc!2SrsSKih5x0VIMY4yZ(g?x4)ltLGrf3EWizbxB<X93g`XKv`N+'
    '`=1FXLkxE!#30ZV?SKd<H?YCItNhcexlZW!!VbNvcF<^nvjU@3_epAA7lq|;L!w-kHV0m%uO|dtuR!h4(;`P6zvFc5>NF-*Tbg5&'
    '-pE%vAS-P(t4|TJ9TdT&Yzi}~^4yndOp9i9Q1vD_wV9c>xN0;jJMJfhR-2jN8dz7{3;cYEYC!7(_Nst+2nOL4h@VwJST&yclvLB1'
    'XaQTqIF9wwy@CoFlj<AdtCdq(&0%<laJodk#Rwu%864G;WrEf7e$ef)=5`SM#z!?XS#81|bU((ckx2lp1P?=#!-rtRs$a!wTo#TN'
    'gGU-Ue@@Ly=0q&(WX1>^xKdq`531mIueM}RZlz2gYLH^d!Z)u>m5U!>U^ye;)(8C?8l!54h@w{#W`43)6+M9F1!qFA!qL4TP_}`;'
    'RUSZ@Xsjlv(Z+a~7^SZ&t8)Bawh=`4RsqXmeNY}*CG<g0t5aht2|~uIA{4Qlm<_Wga8PZq3-VPPCg^O{iXwZlQ`S4EP=#VLD97&-'
    '^or7s<W9A+Y}7`AZfQ%z9(F1?M8BL0DoGNHi=BosjmA&Brm;|eaZMs2omBKwFht~0vriSTB2HF;FHh1qLXhARd#u@91wUZd3w04)'
    'o5_T1SJ)lM61AKNfLBZa8L2BU@dHf9MCUtbE=AyQRn}w*PB<y(<~122s_9^zg@+7z5xzSWJ*2}46myfU>~yO3Eq5og5-}OvzA;fa'
    '2a+EQfhgh?qE{&rVBjz06J`zEHIfX$o{K2S5CT^OCYk)3hqHtC61hU`t|MA>pqd?SI`}RVp^EdQ6EzxX8@MZG+e3!uL(s+)=9qD+'
    'U13#&!nLDdBY3bF89`vqevBgi7UU!t)A3?u5p3fTa=c~=u<yJ>e2={}9GDD}BmwiA@Fv}YmpHm=|A|10n7RW3yK*YBu88Z_!*$C<'
    '75E_6Kz!g&KmfAv)O4(-#~NS{S$#@Dyc-x0_sI-NXXxiU+-o2^Uc&r*m)RB4sW2qFsk|gO?Tk3Kk~oicxDVyL;+T~k?=pO{J6tDo'
    'd;p^cvn_jp^AvA}^8@-h2YLe(9jM}#=+T54OlD9?;(`N&Sy0Mp;6p?FM)QTTdN?ox{O*#yg7`5h1Mnn4)dux~r*RO+)&vNgr|wjA'
    'A*O5Wp8qJ4E+MKldl@l%1WG`)NapuOJo|`g@qR=@WkCCwrbuTJKR;AgYaC**65Hh5HP9*hzOse)4?3ysEE10-ut)%(0*g$h!VWrc'
    'p(0*V-Wc&kjE#^PNooyUcMxW-)VLMlB^Ak%5^M^AK#+?&7!fF&v=1LjfenHa)(avz9zxj2oMBMoLZn7WA9gH~U=XyjoOel=3Zy1x'
    '&64w!a1Kk<D37h2G~&o}5>n?|l%``3HQ(G3#my+SF;nxEIkRvhj01eQL`AK@gP2p(c`d53WnBXEWvl0hV<cADCN5NRoQc#5uaOZ?'
    'b|NwdL~SGLk!e^W5Lkg2?q*r`6xCmB$Kn-xVt6}=#_n)4&w7M?b`onhs*6~5-V2NMNJ|N)60b`VvQ<7H`_aWZ@CaMkufsc6WU)p)'
    'j_P$bq@0|J1Hytoc2;Q03fmDD!QnHGWr2!tagxYQH!n#(N^WecbVd#zZhR1LREEKHz;5lCgBer+r!wo^MKukULQcKN95iqu^D@>W'
    'Q3?2aiMfG~TF{Nr_3$<v$3kTZEc|oUpdx&-CUaCwGz%Lg=_zV5A{NJPSSLxU%o_QaLgGc3{lc%sk4n@^;>^u&XaJ1V8Qu!GQb~~!'
    '*E>@N`<|FKw+Vl<l-WW`kwD{f$Offn|F&uFp(w|iH%XGLOx{rTG3={GWK@W7{6*RrQN~sGV-1rT*;X4_;d@v`X4~?>C?QKRd-W>O'
    '-a3(q3@EY(ZB&s%+so0)R|Z+3hC_G@YNZomnrELj&g8w#U>cUGd&>x*BLEljv<-EMB-qNbkd8T^+M)?I;#NRw#P_ODLoi!?APM;x'
    'YAcZ{=rx=~E95h?<rlV|KdcX_o(F-U_hoWv0aoh9iY~^caE9oCGeL(c9T?X)y`BNJRY1YW+A0VwtBG5IjTiT@E*Ad4tx!lty2V~_'
    '(1~{qAT*F%>NJZ9t}M_?RKJp-BX;K2WeQNC=L$=FM)^~WEJ|6D82M~6J;soj;KfVtNclh@AJj;!DoaOs$EkHmsngroMQFHtQLl(Y'
    '!=6LzBYDV6kfOI;gaWwN)mT|g#tc>xJVAG!j0nU}`E!&sp)=vsDX|eVF|rt)&Ukr}wNOiF`7)@fgYqsc$vN3Y2VKnzzu4=`Tx3Q8'
    '2C=5BghWjpOw@yHDd>bAn+R6A2}m{R5M<-6FJjtiQ*WiCz#>I?k{P_rp!+4RlH{}6nBAO;y(6HBopi($p_7h+rmT<=IFU5&gbe=o'
    'c5|*d`t2SKXd<&69$DmbhLji@AP){>bfp*#?QlTe$>%$N&;GWMOr;eL(q(WTrj^V|;6_P>RMsvb9st^46}=#~#l(S@U2#es{76{$'
    ';`q`O7k+W#pbs2m0Ai{oDk)$D1O<6nHfgjD*yg!N-zQvLW*s~`^=26o8%U_2Hzx*}$Z$}dNDdpEQYA>%&mrlw5@(_2v3jDiB=cQ>'
    '4~XZ%K`B65Amc7YP79<Zq`Cy8ZorKdg!pm=1oFce!zcE%@i1N|S*oGBe3(j{`g}MGHaPY1@B{L_Ao&ZY3>fw7uq0%`3ff{LLKr_n'
    '6=Z(^xW~Fni+vymAr~a5<B(^T7VsS5d~#D#i>*3a94mU&(?aij^rmQxUknb()a4w(8Z*F%5H2_?phF~!!FfWMU^#G5!l21qR6mK~'
    '2t(1JLe51~XaXV!hzVE;SC`YIr-bKn5StggTIf1^p=82{fkzFIbxva8r}HFesJX6OTL|XhzMdHe=jyT2q>~{9z$j>`9Eqn+aG8PY'
    '&?Bkku5ulI>)}#C7N-G+u)=6Cu8CfG3!-&hQf9(0160q0l1U`%*D1^3gs0pI^14QOwEgppeI*V2T7i(&$&zv47ROuMLT|)@Taf(W'
    ';#R403-2nsV3V!OMY#nl%}Ke1!Nna&br=i;U=P3R9-0RktCd^utAI;7;1(?GxPX;#3$*M)l_hv=yh2iv*y5}L1@Du|Mp*?v6a7|R'
    '!8)9+tb%t1AOEFx%pIaXz!03ggY33zDRi*=z$dIGU?BO8`L7Y4fdyn?8XbFwm1%fU<u4FTF!j=hm9qw>!7do4pUc#?p*+Kj7P_zd'
    '>x`*kPa_go3|GTULpcZ%-+d~|K?oPf$q=ux)5IdVngG`&a*@O?0{6g?h{`?qLWMWJ20D{RGUQPgN=GRu^B^V-1s=?U;Y*s8Rpx=E'
    '7pS^}QRO;Y<sQ6T6W}g4Vl4Qq?89V-rGtI&n0)zK*#|!o{aa-poP!dSePCs#tL!6#EewJ>(GCXPzq2)OpX{)ryaOXA2F-;yZ;SE{'
    'rfZud*rZTTG1nxW95XoBx!_s(2&Wv2^s0vmr}S9qF3IQ$X^TvRLj_^5lBB|P#=h31JcW-9DYr>^3aGg=%2FIVcC&C4A*O%`k)By!'
    'LQ$5&{fONdWhsn^h#1elPSSO@vJ?zundHs^6*(_feM-tyxDl69y>%BCe1zmcGH^sQA0cKBngSz{S0ZES@g(7JgOkv!n2=ZrPQs4`'
    '9x>)5=x3Fa@I#kGie_!;SXbVc+7cliuo66XxY$)>F^K*^A&{KPgN%`=6T+0dIhg{)$!j|y`a!CJ(Y16C5SQAsL-rTHJ0TNVF8j$`'
    'EiISlQl4^R1Zm(YC!~9_@H?cs2-QT`!DyAMoES;fy%53Glo0}7%*t0zScx{mSFk1so}qS+%2@<FHruL>bQ)yl#Q+5vLdld0SW}tH'
    'iNRem@RpO7!N<@s7G5H77h01D^DI4l)lD$75tT+qU_<{{v73?TEL+K^rJBcB1Wpdv3t-fR8VEX32gbovR5+xeF%D0VE<MQ7GyDT!'
    'E<zEawkq9~BFtX^;;W&DmZhcxGK782ci^p+8pL4$v}-{;m#9mjQU}vG^OqB+d|amuwj^aRXlO-0<q^H)ZK|q*1P%rw)0Z>KV$jDk'
    'UprRJCRj%b?pQ*s*$8MGA;Lj+yvVAy2tmi1Y9Kw4sE;WnC;m_>fr2jgdYXwqVhz?{93zHmspzPNj<l70hOSbBXEdJ+WW4g26S+`x'
    'sK?O6c<WZbsw$2%y-Q;;V(ID|jy@@eG4N5Us^W+xmq}G|{5@A&RTXD^Y^n-9cwyyPAHi3PK7@itW7@N8hK*_LMXa2Iju1?jV$)Z6'
    'w$Rv1lq6==SE#M$1A3X;@byN6(==ZhJyh-@`xovKkuogIg~eXeSa=DIxpWA_iHil~ndE1$ti`~kDV(K)$#Yqu6olKMYnQ9Sf@SKx'
    '5yePoDvKiyDfCq#C^AeQLydFG?Mv_!r}OiHB7G%Rg~gpt+6v(>APq<QGRTf~r;_A}<wsDu>MD+DE;L;QXh9aN_-tZRR2T_WPm%ct'
    'Q!1oYNpW0JCNw35sl-=4LJg7MP-!F>*>S2%6*f`;!a?0fjfZ#%iyZ2}!Si4jV#K?qqi~V8ZqiYRzEf3EPJ-M}I!`KsD+l2a<)oz0'
    'nvsWboT$__C52<S;w_AGx)csV#;R*N3Paq-rlT;FMTfS+Hv#g}ra(}>p4Lvxs|r$|5z4QW5-fZVbsAShJayAa49Y?}1tky`a#K@y'
    '##n9@08uq1gboohF?VciN=k@9gs&u<L&6Bt=C734Nk<7pfYwwUCHUqW;aA1w)N~Y57=%`ewnSpNswiBPzHyLflz|ukIBNr9^rdMh'
    '^c+)VAAF?^$6UB~2o6H)vh~CT2~eJdI>q1-qXqNG<rX7R{-oT)@IAw(n$S3*+B=e}1cbq>CO#4|JoIi;RZDV$PJLO&a7`;Q;@scY'
    'Rd(V@z|k?V3q>x5=%C-LN)iaJElDP6#x$J0!b8D5IMR};yCfjxT+(0y*ox{SAsR0QDPoJfAL~LBwu3En>amG3zwLxqcNHwwR98um'
    'GkmBh58wgdA#evu+GSQhmTV7MDXNL^Lvf9RSZP;DB+5-6#$%Mr&o%S1`bez94O#J#IPCk-MlKAJ3;5p6q=gW8y)#V<31sG!s)d*$'
    'nFQy+Z>&lPXOyShl#qa$)l~`MyOlD$!npG(HzkCi2j@Om2-ghYxl{?Evt;e0gizBcO|9H^h_PxROw%6nq#jo*(F%oVSG;RVNC;6f'
    '`RXjm2s$hN45;lK!lhIWA_qm45dIqVtvW%XgLHfmzN#Jk5G&6(=|tVMa*ax^t&7&tb#k3<pR`l8YUlyNO;nDKi>NdmgWGjw(ko1J'
    'd~|%x$yG)Wz&u2a;5W@IvC!eF8+1|@L<h%l7vI=Hx+c}`p(={ji+!2XmM*8h(kAX*mG&K9IH)d8*>ANVBpqHX2$67NKv=o|ZFPz&'
    '>`&wdn}~Y`PpQ(L-FLYPb$;V(P)inH4<vh+eU<55=TDn)ATCI;4np(z3r3nq2TI~3(M=WTn55bTR)Hu0E!F41YxGs1Tg6OJ%(y1R'
    'JKS5YGQ49WE9m*NlDZ>uJHQOzE9I%fQXyk>X|cX@dc?|UtkE-UCb=q<<r4Q^X;&^+?G*%iYeDfRTG?KK8@0@bz;N;gFhsn&RWwUZ'
    'NhGF3`d_Gpe|#6@0y|CsCoXlB=(A^rAhTIG^68~*hBsN$gpMpu9q|<tyXtCy1Y_CY>m}4e^4H4+U`*^tLy>EHcuU+zQxg&0>o~i_'
    'hw6(bd6be|sS&-iOU`Tj1^EudWKK0a0o9ST0=}SEBUgqL-qQ&G59o~x5Jul*fDeu*o9HeEgima9gM06sUNU44Zt1vMU<<DiRIL;V'
    '0tr~MRyZZ4;hKXbK|~;DqLBvRifzFvnvEG$s7#W-v2ofV4-trU9YejvnviOcX~VZFp8>mk@Vcru@r@^63p|A^fiw71y|})h`+B(>'
    'qA#+mZXBnV8~5UZ_0S5z@HDDu<!4_VyWFA0d?fTZw&s$0oP8w@^vsnqa-)hcQAD1ofm*}K4^nUvypyPmWb{Z-#(J5_K|bY1D>jry'
    'C!9)er+AE!DWX75md`z$X_Uir%bhEW^o2??R>PaTRPHS5c8~{Kb541mp(dRgN!kEs^2W$e;5kKp=j_bnTfC7+0_)_J2Ro4)j~cTQ'
    '8)Hq7pxY=_X!Aq%wP=M5yjRMwrAlbfOm#%^Zi}fA3*TstXH5Q9MZLY;X--wA4ovd41lKa_+2QSytAdgUT+*5(je)bcyC7CR(h0b3'
    '4HvSm{KR*p%t@lO=*4^iG^z&7*PK%|nJow8a#Lk3@m$WAtrl)X;)l&(!v<-@oYo`GKu)gB9an`h)$WXhs3KZ5NPrV7le1|Gvw5$U'
    'KjJ}Tv(bS>#1=UVL?qdy9II2DL)ffpSFbpMnE19!H891cF-!?T9KIfUaZ03r#p=6>Y=6ffps~WWPZe|mOVXWP7jIQI6n>wyR0<na'
    'CbonGIO=e{4tHCW;6=;iq?Bq?F<_DN3-OS+=tX50>~g~;HIWv4-;Y@qyfWMn_h1BZr#7)nMXd}yLYgG34jHtlawf5WGBwc5)vJXL'
    'Wd**7#6xzMms7+;n8tDR?DHC9h<di5;BszY6;CuJa#>Kk%t>d++Y%=sB_|f9Q}8J@=0ub$Y)}^;Lv^*NFtuWh86!b7feGSkBF}7@'
    'ak13JiH&*XX<{^FkuoCmOrAIyC-?!x!pc%8?Fu}9j;b*>hL|f&9pha~kOxvUmPtf%4UbX<B&YZ`>1dptL}NX}LB{T4R)x&_mT<w-'
    'l&(roYXrVve9;s?aX#8#V|~Us;~MkhvEfnND^W&VeLYmb%7s@@h!0X+M?|xWd6LVVl-(4w%yFF-<l$IHtcsZ_oeC%0X$qr6GnM8b'
    'Fu_Opi<mV%1!P;A^Mpgca<53?lzhEH#hJ!(1yIN5Ov<Xo@hG-hI<CwlJr`~>_EuR^^|yMPwOoLB1jTR@7Y<F^B08E?UOle#Ce1vq'
    'CB=5OS;Wsgt|fbyo5$4~&abI1S85qt3+dn)o70d^#BCNBH$mbwSz8fLbj{~V3@eC6&e}OAGP%j=$|*a%zf0|Ea7j;tI-K|8n%^~s'
    'Dj`n12nDGAP7ADw>y`6}sy&63G}mhszRZoHI6vWhm0~JlW_h`0FYZ^qaNt_Rm{x|cq}X6wQlQ0*8J@sPwaH?fqQ>|PLG!@UX|u5^'
    'I44}61>sQ%%y4>V4IG~&t%C(CCS;hX8r&iCJxtEnT;vvvj?vBhv9xAb<l>KYT1-kkvizN8mzzhHHj!L*Q+!9Ns2mzuNV<%AG6fQ{'
    '2M41{U9<S2NMjp(16Lv%6*49@-z?YQ7D+9*?xqA3!fBC;9)LwwqY<!(on{2q2~fZkVIB?#&78E}qAXQ)(uQIyCppQqJzKrBS)6k)'
    '&%8H??d{v?^K6kHD|Lo0O>Wwpy|c|tO8^C`#Z8;k4I~a!-#L`i+&7~4U9}QUXjr(-4rM}wogB5vy4d`*o<9jwnxB^FDaG9Uw4RI!'
    '*Vv67U#EUr&cdx4ALFoT6*!?O<y*y%#msQ(O}CP{n{$otXut*RhF_HGi{*p?S5wpYP16pX2J33zq~^IgNn;l7&`F_DF*N8!tS(dp'
    'Ny-#=X|a~-u4}VbJCq{<R!=tnUGqKQh%tIVSs%33%gM_Xj;g^LakNBrwA7n_7@)d2;T7`rZ1TcvH9>gu&2hqpUiD@IU{wjlJT?9S'
    '7`Yby=tA?A(1QqfS#*(1QGtHxni=9w6#xn{o9ZFr9hiEE*b7|`5f31;A}BnZ9-2DzsHibri-$`6LorgU-oDgLr%?Tl7`aMb#0he>'
    'SeTZoq&<-Y!@Uw2F3#@jKE=W+I*n<827$Af_yi|+t3bE<3Ay5jf%=doEL|;TzQhu{COcix6h|v}$@v>c)p5vGk5K<j9EYbS<mx%3'
    'S5fK+g4nrkQYWmx)TU|XOXPCMxWYM#?@R6?g4-4GdP)%?PK7Pa&&XGI%M`K9ff^%vExsk99ggv-{A9F7XKfRp#!Jo^BxqF6AEr#F'
    ';^f5xn^9A0L^46pd(bT>U-LBMX-l~xMl4YPNKvlZTGOs*KT#rwB1Z#hTt@RX0fr7m+*FcwQoLoa>kq}GmLy@4W(L8f=Glc&9h_Y9'
    'ursNhR?B7DYC!X(#C+P)BvML}1Oib?X2wjGvPZ^MY079wQzF($39^r)E<BR<Uvsfe*zQpeT~thH48BvS(Z<^$=MUkioRp1rUCls6'
    's4itBp{Y#e;&{-}!~|gu(}|}&NM|}Dl8M?$Zi1OSq+Ej6TifKS;Y=6b1S^6Fl_#K3U#UzXaJG1helo*>C>PPn8Q@h(Cydb)U1bx5'
    'nUYMJgsWlo$f<NqMH2GF`$^g+oSgf5LQRJKeN^;wM8PO8A#{YIs<nbRNo!#!F(uJNAY?f=H>4bz+m&y|O36`~nv>syrhy7o<3`0^'
    '7(ytqYvIriiCD2J_3ZE{EoqC6ed$V`e6w8V23mrMS)=-dAU4az$U{BNWQv4x>!SfGX>r$Eb3i4KCsNlNaZJNZ?U2YKQrEMS*{N2~'
    'H=b&Wr$bS}nh?%!NJnZOK5m$cCVAE*03-xMzG#1IB|IiBR9af9rPx8Mt~mo8ikD<zK-@G1U&9|LUm9!NQz|)E8m-JG%7s<<3VGgY'
    's7ojl{=tB(qzk!=RwcSGaH?PG-eb&FIZqPF*z0<HlrUX~1b~_w;ZQrYRR93^qigA$Zc67Y;-eRvy))|Lm;0g3VmDEusv`KApG(b9'
    '$taf%9CeF=Gss6gg~~Lr|I*bQE54)I^1m6*MW|VfM}ZVpOWmph&dub0f#?SO%KNH-OR+jt6!?v)u@u9ZrB&a~T3HX|XUycLbbMRc'
    'Xws(5xQr%R1M!$-Vy=3FhySS5aHdl`=Nukm&BCx!!n);uGp-8Mg`?77u(9SuVCvlRza1QmrE&uP{=2B4MqwPFQc12R-!*q~?1i#Z'
    'Q#M?d1U&&Ay)#|at@&MtVDw&aL&wI|oCm8JX>uh_^JlG$LGst%`t}FTccL{w9UUb))^3EPQ+8c72i*zZLC$)_`DjJSfMVvIxZH@I'
    '_r#iNQbmMcv;^<OF2#C-jhh9@2r@<<ttnat<wUElLm{9oVNF?D#T;rlh?nl+4V3fGU%rJJ00zLt=!i+H1<!5MWWaOe#%reci3R~N'
    'nZ)ZgEwig$9p(o%b@F|JgYGh_>fujYgozd_1$>}-c)2#Efl{tRTQ!g3WmA_rRA(fZiy^qusj>+`$xT{^=rwzZ&W0o$1bl+;k*k7T'
    'AnELBs5*z|`le0kaGYn+k~2-_+IBS^;)J74Rq5c6uZ8y>qqr=i9Cgj43N#CzFQDYAaU}k{LH3n-2&QZH_X+3BSKh&@UT2!F!F3dx'
    'S&`)?8JGt<!}!~5$~@Rv-J$Z0s2Sz7`nh>7Cj~nN%_Q$bv)v8Kl~omky)$MC$MT)8ShF;F1`o$6bc~2(#Jy5J@mky=R4kALM(Ui9'
    'U{{0$p<#I8ubQaoJ4t-s>#?&)Ge0VEqQ2o?;of<a7dupNlJ3_MHEpPLh$}}V)(kRn_Lb;^RGX7YofFMTk6c8lvb~;fZ=9>Sn!M%g'
    'Obx>@J;~h+oyBvSgz_AKIYdp8I1!ar)rziO%F+WGU3DMohlS*ZEQf$ep;rH)BhC)VZM9~G3*vxrsDjQxo|Zk4EXv<Yu@%O7VG}wU'
    'HT1cZr*V}5@trN3ZKyGV$hG%Nu*!wFF@Nm4#*6qyepm0L`4}g^fOA3w2dx(BnDx{^lHm$<QjZ-=w3a{deQLuLSmIA?S8IsNI+B`d'
    'YPBW)@UvVJXLjU*-YCPgk)y`-?g`mjeBI2|XaXmsdkblI(5)w@ensEur4Elvgq_?p!aWq<$wwIEDhd%#o}7!8OfHgyQQeCipX9Y>'
    '<xZ@t$`?}h0th;7Hs%(2ko<G;FyiYRsfc%}pwK;Ay<Lq!ke*xk!#FjCJQ6Ab6k)lc;sO`mshld7OyXnGYe)$;QbbxTB-GJUw=>Ff'
    'TPHPwkFpKJ3B11DS0S26m|Ct}vyw$-QN6Rm!E)P~mNmk+5JOgkYHd0)$82c#R&4O6uG&c^7B=R(n$*S7vEU(P_gWs7FkW0u4fRn<'
    ';~$g=vwXHeF0lB_?$?@TP6ldD^#TSDWxG(*!2zY9P&(8)w9djd#C4Vw_0F#zBog6Ll}8Z9_c!q^0HtF3KcQk=Ycbrp`Xl*rFB!pI'
    'wU`2XWD4r{9EHT0>Dy{Bak^2E!~npXwInj2TH+Uf03e1I{z&q%fmkoOa2aJ{>#l~9J!^<ja(qffmvBK1zLUmgYJ8a;?<D<0b=XAy'
    '%#+3a(7n!PUGcfzb^%xLKI#oRl90mF<;Dx=X`IG(`I>Jh60MnA-$>Ic07Nrzw5dMj$xinz;OV9Gx?Zs*#2IC1TwAE1?ictO$wXI^'
    'Na+=S1-Vpi#y&~hPT$L&a*`-{ssxaW4fO*7Ub49)>gkp8r9hd*TzaQA)|3^w$i!EKYD$GgDf{Z_WMS=;F{O?R-4<m_d?6s#ASW;1'
    'tLbFJEXEnD(-*<tqMvAC=;x$<0S$Ae(`F~xDYQvSu-qj8Mo;-0<L%z84@N914Q?i_DlgGZ8>aw<A??aerNl`bnkmDKtIT$`rI?XX'
    'kWbrv*^u>G2N+0XcbBnKAq9L;OG2D{yBivDMOg{4W}R`Vj2f{Zwb@J;X4&SD02EjTm#+769YpKEbglRR8O>dGE!3H$*pS3JImbv4'
    'FJI@-96i~)(Y+t)45<s#UD@b;RdcwePz!cud}vMiK;fuI39aM=@gH%fDgZrEYpwRcKytWoJPMt77u%7i)0|9lHDQU<8yWaDVJSj}'
    '>Tzdxe`#tAO@K@uE!R*uTcD~jOc?5W(Mh5&F0@(;Iftscc-x)YR#KtZ811dk8;a|eoe>#{*GOjwxaWlPP2H2pRcqmVD_g^!mbFli'
    'prph-`CCf{@isVbGr_5}v`=>Vr07~&1G}k+wyOJ<8z(qE#eK`gh%=n#`XP`|W1m^djUA+nN=^n9*~ERzSD}+?@N6fC$fvGFKlv5{'
    '871}33HOd>mi1twghs?UR%-Z@3nSJW^{EEngek<pz|!0q4Cu6$v6<bmH>~^>RaB<St4=D*XT8qG-TBs9X-uNXwVu5$mSJ7LJ?2Di'
    '@2!@h;v3C1uEGqBqx$vJ1svs_639_sA;u{!LR2kb9BkCg@*?=F3Ax5T3zvLxQ(^@4<D)LT7&YFbz?n26iL!JGm!t_ua?`$uW35sG'
    'AE?}XEmc94j{#&&QI=eKzW-dBYnGpLi8YnncUmYzDTs^_Uz5a$drJ+mL3q3Epbl!f1Xt#38Rgl8sW=5T6OqvXH%!XY%@m6Q{lQg7'
    '4Eu9~RGyKm_<Y$Cu6?*-k_Z~{#dJ<SuCT8MkaYiCJ+mmJI=iA$G+F3S!OQ8~mEok!$p|eclHT$q_Lfpk8co*{D;N{BrkWDVK#0C-'
    '6g+LnuH9@ia6AWBi@4xWI($bCUK#yZjJRcT)T^U%4oOcTz5dn*Q>A<hU&(?xUoXwpAQ6AGES4!sWF*lXtcKndNgP3$7|~&+QVW;x'
    'M!JJk5TaKm)gTZ*#v&h@!$pI3XU&Q^Cj(d)eD%Jdi!(KSg7au<^aq2()pMD(VQfitC&`O8CA*r^(9NVm7`6*)sT-Qrw_J&mo2#W@'
    '7`*};sy%bUmCsBI@kl`SQWq6-sEWl2K0a4#4>2&x><VnDRt#WmTWJ_pI&7)z;)F9WTbcyE_Sd{DC^$9Ndly{`#Ac0eTLB@dt<7dR'
    ';X0X`BY|xj-|nng-HFuT;W)8PHAez%-7KV|WxEcFc2aW0ptQJ+Yo<%Eu1rRXEGE|^QtoH{tUyPL<X}1*>*ckRkkwm<YpU3{B$6tb'
    '>tHZ{>!sJIe8KabMO(9~=qXxX6^&Tnn1d(9O&_A;*de6MUQY?}prB?4mL@lO6~3fCwuCEgn{B&o;*`j&f{?G0)$%|b1xvYISmR{O'
    'RyK<1QRqYG&%%Ech{h#bljZTKjp62}lWt1QgAlUosv9N9SCxRJjPHjW^qi4xl;f>0Cxb(<?#NNJ>CS6S!Qx!-&bHKRruA6EGn~5A'
    '$#r<EpH1{q!ada?RTJS+g4U^)?{E=ls~#dPyPAT9M-=eV6s&G`>ouzu!i}^x2(o+mvky``UoR+%7X&=S^0juP1z&2<E}6zdiljN!'
    'TlIdr%Hm>k-{yKL3bzy08Fxr^9gU|<kz<eA4a!nniyEiNIyqt4)CotCMOD`bpI%i#6gGL{C(cb8NQzp;U7)CCEN~u!X(^cqPgbMl'
    'B<QD&#aD^XxV1<z7i6sw9$W$t62~A>207MEOGOkg6gQ=wGpwCBHDO$EOWa^qexlzNJ17+p$=NjBsc$Msh0Q@yqWP;KHQ|b%ye2ku'
    '(@p3~IX=XyeVk-7U1dA5(aIB`5T$yNcx+t@MeF#a-{^g|6S<qVS90fj^l?%=xEx<>4qGqZpjIwpTRgaPQzbnfJvpQI>IIN6BBanI'
    '=c2SLZD<>FHF7wvErq4T??fD~IyJbZLAs4H8J+S_DuuQvFy9DLzE)w2k}ro!L#0|cyqc<o$eE0(-KfuvzsP5E4)qjr9{l_R>|*P1'
    'EC7`lM=6|TT81aSp43%a=OkN=cI!89N(~Gdu9lYTRS#6Lt>1s)SDJa2<xXaVr3f1*;A1q^lfjP`KFG1OA|1k8<lBC*aJWy7)=hVX'
    'v}(S@hTF9UEQAYD+FAmL6PIIbyoEpR?A%l?tmSs*=7~K?M(fIWNulTjo^_M4RS&ab8=;w2It6>tB<7*QFEj?jMeMBFIia*qmCFbb'
    '<umqmzq8nomqK)MPnSQK>RZ}3yC$=r?%Wltah9l57jrAz!AESmdFR+uoZ98PC>p&OIjs37xx!w@G3su`y7BoHy>M@k`YI+5`*VtG'
    ')qzlvz?gPDkMqH+n-x@%O5-)`CTjVezIeUVt-edsbu~i7lf=|2czDS`ZnTsdvtd@6s<Y&To%YGJ8Xggu3tGd&smXh&<BQ}Lb|clL'
    '_Lcq-IIox}&e!XEM1_)^f374Tu6)G4T-1YGPag8fj+H(&5e~j)K44w5aqw;J=@vd=(5h1c;N|jXHMx!HKIe7|jHIZPZ$8vBm#n0z'
    '>KQk&YZeb*--@A?3*yT&PLktv8k3v#!!U+QajwvDjq!NYLqD_=N`|^Ho)^b=Pn_q8^E`2$C(iSgInUu@x~#^v6H=Gk9e%D4?R*Q_'
    '2->aQm<&lBxW%bKDj)20S~yKEnEB*uX;XXZYCW;S*pI2*)9<u4L&E0=VrPBxwc69<r<<%T_{){->(S?F!gPI1-8T-OENyo~_~R_O'
    'MQ1zql*f8f3)vT`-F2pgO(`OU^L~8$)@rAPVfE98M0B~`c$RzS2L*$-v1Wy7fyhAyuQf5ed?gOKvIm7$$e~3t_)2|qCh%7uW4j`s'
    '<%3XqMV?Eng-Yr~I-%+DOp>Om@$63k-;$VQVGm!WUc-eQyP8H(f_?{GF0S&otaG~?Z&5))sgph4rPTT~g}+N&D#=;7!E;c2a}J9d'
    '@fVY<i%7%p^)3U)NJ!u`0cZJZ1U2`8FJ*e~r$C{@SJC!Kz6_IM3llo6-j(H>YIM44jsrK8k`1U071~v5UWh9a-wJZ;hrc0WTQHEw'
    '-8QnM*U`)8qQ!X^W<?~N@pej_%J@#Ly7PGYVkUdrowlnGCr<aO<}`nGteVZ7ITLKFlgpLpyZC?Ddy}2XvLk8iRRn4q5jyj*BrEe9'
    ')U3)xv58i9p%wxp0RjXFkT%lde|Nc=`5BVr3kWC>%!rCF;@)%5;G?;D*+g_HjCSx&qm~0Bcw!5TO&!$`+>TqZL@+pqlkiJoGnQf*'
    'LCq#S!&gVlyKoV3xauSenq;_gbHk>yG6xGb?R8zB>S$#u(plZrxNDBG)dVM~9*((bSL*Gcq*b-={_yZ5sva#NksWtB509#H60wKZ'
    'puSt0PxD#B)(&H~=0GXc{RA=ffA|hHnONpd>ZROb7(1zl5K-M)ptllUb!YLaZMI%^K@o7OD{}{7HcMYL6LuT$hVCRi@Udk}Z;)Gy'
    'c}iQg_LLpX`0L;Qa`>$Q;vNoG&LN&^W2C8iTdhk`e%4RHacbVLMU@r6HAJUmVPg(qBqD`INSogp0pIBC)(AlCXc#(&mg;7sr!Q22'
    'KJqR(wlB4?xvf!A(pE6N*oOA;5d#7f<5Z>Bp-einZ!_$w<J8)rI*2I83i}l1TAJ7^KJ`$1CZ^vMv{4=Vj7ow0!s_>M$TCdrYCzxg'
    '^P<uVwcGI@eRv-eXGbogscvbf>Cv(+ZgVgq#n^YHF>qe~pPbz2HMOp4(2=%za2&)8PEG;zTv*?XYQ)_~-oTuYRk6&2IIG{`)8X5!'
    'N&EB_8*1A|yh&M8%&D}~IZu^vvfHVKT*?R-nR4p+m^*EYV;Za-nmAZqImVMZk3Bk4TQt*V5vZ2JwNs9HMUlWyADynFi}|{27U_vZ'
    '{>ocnWjBd5o2SRJfW&}L_q<4?VKbz*vQoWyuK;dFupDGB2Kd|iUG)n^o7#>zx5Kh%e$L!#%h}Yu$~W1{?qs)Jf(Lp1X*ef53mFsQ'
    '|6^78;+tZ3(051g6pRYn_RSdIB#3nd1mpTH@daKg9Oh>;z{w1yqTV#@G3#M!CDhJi$_wZ6R3i~7wwlzwMMb<0iPva07sSf&$k~dr'
    'v^6Q})cnn7L4>+ec0=F$@L`RbFM9ZP$em3E_Oa7Xxu_F^m+9u}R-n|kl&;jC$(GkvdhCHTccE@FAtrsJtH>|;?xUO($zGak(>LlN'
    'ON1xBkKq$6C-D$FC$O1rcBNEZwR7BPFOYU=F2HzoJ5cDCSI3SN3(Vq`xpu=45B+KrWI?L2q8(b^E>u`poh;;Uc9%`J7=hrlU!0RR'
    't~?Whwc9i5u{jHpZ+%tmd|!(+#-VRE72nIIT*5P?bSRork4Fi%;e9yf+C<VkIz@eilV-+1(^iRC{yO_1WwtS}vg>DmjCRyzb(hf('
    'fnC&(^S}N2ALlq{lsoQT5Pc_<<y2zmIjQ-f&L`jz8;F8G)n!)lzheC7ZVn!w;IGljew|F9q7FA%M$yKu4xF#@&$*qjsn)Z%#QLZ7'
    'O%iHA#-zt;ioEbu6VE|!F0u??sqOR&2`qI+y0b10CsJO`YwQBvSzgtiv0jYAs|=Q?>%O-w#v6CtTYatEB^2fBxX1wJC(knKH@mtT'
    'AhodsxBXXeLNhM4M=jF&l{*=~eEYYnw|fs$h25$bFQ?wLshC7q>n!o6TDn@I`t+)%)bt83p$@q1v~r*5a&U<J@t>2aK=r`2Z^mU9'
    ';lDCzZ&l>ADoxb}?{jFVZM4N_haV^)A+IEXBx_Uoqa}9?Ve^00UDXvG@uz#{J13ax4xga_vz670wzEgAioD`-r&0W~ia4ThF`fuS'
    '^RtCON}`$g(Bsc|h;P-4&<jOtEM5k$nKApk@0yKP$)PXq%H|5Cs+r)D<8YW#JZyN|7oY47w6TgM$OWbZjnQk6nnqV9$>iC7{6$vt'
    'vV&mBt?acDeXJN5ngnbK9%FPgbkQn=w(GvRWObExWxrt51xIMhzt4)dd7`fmF=}|>ciN&mRY~DbUE3Bso^sDpno@7hZD-d%#Tc;x'
    'P_e~1$iJ?}s5TEVY;z0F?<fnYsM~Cx2G-xMxZp3vK^y%f<gtsoo~X16b@j@&8B{^0*K(SLzRfXSXoS3LIXln_eHU|r4?C_#*E~jT'
    '6SAd#iZ?m^W(BU;!rMh?YhFZJwVN<fL@&-ib<odl3{hK2WHFJCk;@^qHIsu?v?5Y)Zi)4$8Fnlk9CAg_%(=cf>L#@vg@qVyzKwd!'
    'x4vj??H6rszCR-2H^Ru|u$f7!E7CwrHKxAJw=tx<-G1Ac0Bs*P;EEa+`1#L4Z5uw<gVh5x>bt0GdM&ahib-OL05i(vo3F36&0h<g'
    'o(KM9#fWW3oT@#VPn0%;?EnGdG3t7G&BMJ?o&M2j<z?`BkURvLnr_w&?|vns)!uX7WVTNPVAD?iqqL=$O<@c4U#pwK7R#1)_TS7N'
    'r^@#gHbZNdY@y5Gw(5F+&BJ4Bo^Idup7_t;;z{$dn<8^xjI!ZGar7%NUh9fr>iN6NjO-CAMl?V7Ek@6WPJh%;50R~TTWuFN*nCtk'
    'F?LfgoPY<e)HaA2Gnr1jgR?abc|@e->YKE-n}L?fCatwSDzJoZ17?D<SZlH0X0MyIwxfgDp{=zEEV*ymT9ce19NOAGTTN}#)(&(='
    '8=Ep84C$M#7Ud)>^7!$GbofI${2?9wkPiP!I{f>`Uv7paQgpo^CViI$uXXfcED~8No`=^=o5vO8lyj3Ufgm(&)(0iWo%6x|BOKng'
    'c)Z5$tZnf1@~&Gz5x{PYfk;K$35n!&Ldu=DeB8!e&^?v0_*CQL5RVVpuEsjOo^CqxB*2wB8|w)!aNlNWLN}Ls?D^<`&L&iH+cZsZ'
    'D0eibT;~!bLpg6eFwpFMQ!>ydU5pGwGp|FqO38e(>nw-A#U<%-FzV)Gc+Id|Rd6O_+Pv1wV&`oRZIijHJKPf8bzT{3rt+PQ0Tn6R'
    '>;~Low9TdP8s(y7%5#=Yod5>YytxuytzLD^slQg^Rt{=xas-t5r_GIkE7nbo0IbE{^oL3oINa=jx6i-+?IJHut*~0iaVNWY4*|1~'
    'H=Y)Bl(B9$V-v$=QMh9dw6yt*V=u=uhVPb3^OOO<>3@v+PpTN6h2+~%PV37<DEQ9dKfrKj9COnjCv!vuXy{({AJ>UUuXH+3_1DF5'
    'kbPJ9*cUBf_8%~=64Gm{IOO&Y=oOt3=OQ)11GUwn?%AC;osSw(Be^3jeFkUNZ;rv-NLQ4BwsPnRQ0BRgipfLo-ovMlC91eXSIOyB'
    'HkWcLu`fI;#?^B=6a?Pf{By0zNV^~phh)r6LXcP1i8RHj_T7Sx=!SecShiD8qb;ycO`iE85RBAdhZ5fOLpm%3Y*o6+m(9y8hwnYY'
    'ponoZFx#MKN}Ge(%#|?aW?@c6x#6`lN*)(0Kyy+~--aWnx|x|xljKcw2LAMRku#&$Xjt76DDTSV1U3ccRRY?a9rcEO%D!BLz#}Oq'
    'qp;&rEFwQwJzUD|CV@x6?Jj>qCh!!$B@x3a>Nx(XEu{TDix5-QHmNkP-Bo>=^InTFe5|CJ8oj=+!kAUFmYh6ZT_hF>P}(`NMGFo#'
    'g}S+UZNRFD@J4!@b8|SMwjdYLyD*4IJiJog<l*ggKuhQ*>INYG=B4`1Fkfcy?zRpE-%Q~bvHZxMyi#2l1H5H7Qk-d5g(VQpEW<1I'
    'L!B{)f@XdJlvn3no|XE^Q`sM9KNVG2i@GbQawxkYvdu*)zE1v>SMwf3w&yOG!h8g;%>2(^uX3qK9zwTQ<_>0$w)PczTZDYdp;}tC'
    'P%-8h;;w~46=F#n^cS{SGL=M7%M=V{>|!Xh0k86>F+$p8GZ_N6{YzNiYUPaXi}67Bw&Vr$KJ8{3v+fOM(qODDrtu$(B){gVS30zj'
    'tw!&hN`s_@&G`J>S!7xh!~dfx-uk-Aj&|S>_v~%weyUf_`OF^NT;H4U*=9J6w9z9p@0NS4THEHo`be#%=sR)_!utb~OX2Vw;P-P^'
    '1jUf+DfyYXTxC7z5nom%k1b(L^_1?+kxt=GO|EJEX;a%IC*{}@_{-{Y@%`iUf+2~|Zk!tqJ84{i9S;l>Jz{R>vX2`;k#!|KK@1V<'
    '&Z>lz4^Qc$DQ($XgOS_gmGW|W)wpo+k~&EUzRg-hzhdZ{l{D45SxIAAM2PaRKE!&_`cmn3ry7>68$3o%4`_v<>1O4+YU|-Ct?)Dm'
    'RLZWMr-i??b5Jd#PfLyIaE^|=d#;V;NjBJWtI8GsQI;0wv@9C2@LC%8DjLC!cj-D7EH6w8vM6zbSdhK(h%PRg=c@L?{lZGi*w=dK'
    'hX%B-tPHNa7_`5AzZ|To*toXHF78AJ-m((uxp}`#Z1SGSz$K#-{}C48%TChje?tZt@Tio-tLHAWw}UwJd8V&-M^mQY%Id6=%Z|rG'
    'j<f1^>&8t^Bic?5U{u#wg?$*x0w9{PJ{P2`c9%a&UM|}Zc)ca0K%>%DM<CN3yP$Ho3(AuJ^&irD^{~L?0X4O%pMX!?HIuf(76wbK'
    'X{&T|sPEL~b}@aHUb`W2V{}7Wv_>~?$RuZf>Z`mdqc&V8iO<#!?XBBvPG%QXo$qVr5C+gnb*Fm9<27%y=2zBsP`M=90ra%#>-uUD'
    '2+jha)zbE}#(Ztta5N`S@`+dk`dVnuDDfImTScXfD74)CP`Ru^CM{Y_M_=1)R?mfL&f?*nL$xnz7eD^b|KorEU;gL+_y0ztar@8z'
    '>wo><@b*tX|M~M5^&bE8mtUWM{`%D$k^ej6l<&h9<TrGywHrPn$eddG-R_g^L4Gqs8qOYAf@ZZ(zr;}abu&{Rz-cCaXmkA=X?-7b'
    'IPqDX_+F#ijF;Aa_XGD^1BI5Iee`L_G2-)#LBRHFcBlYO9);hp_<p1#ZTEv@sKq}2S}-_UQFA0a$`$Zg#WgB(wZO(?jAAB^mRpRR'
    '8u;|R;i_nmXeI-AXF4A3z2GT2;sIvhVdsRz&ovQXJT(6sI5D%wfA_P_r#63VJ@!|u)o<ou4AeWIH+9#i_pF&}D;x&Ci%xQO#__@_'
    'dgC_>*@^%+kGL$w$8VWs0QPUftiN|WTZeUO+83|2M$Qdp9$Z>Btfm!uZy4(Yn8>z0igIHboU@MgDA1FLtKze_fzJl8f{y`EC^tjF'
    'y%=o7znn%BGPvqsgI&ZZIf)oO4EAXo6o)==A-`lzK!jSsOX9QcPwoB?C7iAI%83QvDd?miMtb3+eOWVRk_7)oZpne??Gfb{9dzd0'
    'X93?#o&4JIfxhIaCT=zE9Hc)H3BkjR!MhF&oed4mAA}mHgO|#<yt%9eG&R`BBXaq%+frf|NK}r`$)WQiu>MU%f^VUj#f(49$c;SE'
    'Hz+?~bI!_{`8om@oJE-6F;1C_81ztV+tX1MgL_U~9te>LTN<4-%su5s-*aZuju|V1^X1EKdx`l~P8W|<0)+lStO-Za8dri45ge3S'
    '`Kce2fn(A*u!P+K@=i#Lj(}kOU8asb1f(+lbzz}X3W^$==vUKpYRMtX%>~BVn~vnbItLbLcHt%I`}}l(Th^>DY%Q}Ixuw<=*-WEE'
    'r}HyE!?vI71at4cC#@7|x@@UTl66fnt_*LrKTs7EB;L5u&HsDgA5P-L6BowoOWRwYA7*$>1zb=1F|Md0=mN%Pe27CjNd`Yb+Q`XT'
    'QZSDP@&J4qXrkXL%c)J$PmT~@o6HLV*2&*~T~7mwos_huCSNfBwzDwRcqyF!%;Q2`i5@x}Pv;4>;id7w3q{daoo^W3I1|G*H9qZ<'
    'W%e10qoeLrw5B#2w)$D0HLXPBLS&9pTQxCXVLNqd`QP|WBoQ7RcscHqw}ej%nS~1q5-3P0#n5jG_~lHTJP+!ZbGA(v<78K!jaIZG'
    'rguR*$t|@vPbw5fI=M?mo0ZD?8cqxMzpc6LF;sN$5wxc#S}R3#9nq({CXHMGd#j{S$3+KmKTHF0=W*>P@(<Vot_rY#==o}^+F{`T'
    '?l#wNRl4OtR=>uD&4~vd>MMhOn4f6?jK%ZWM%9I>>+I-PJFg(J=~}Xy{9Kv*!i#UDNHHJSHly2ks8rtCI-X7XeQL{o*`N194W?7y'
    '2e%D8E|9=er4-z*l9#pm+R#*l3N2_B0Cv*Rxn@hYf(;!W8fEC@Ntj?pOAj)B@bkAq{Z6(#{7bRbyFN1;6rIR46ML+DU{g;)E41`w'
    'mbRq4mX%ULMp0;|fVGaJz}?Kk>nLo8TRn6YjgHi|_HU-@*j5IZ-XC~_NVc{b4`yJs1VUkc*nJ1uW~nF?<6QDlGTlO;N4YI?)TXZ7'
    ')HKY#xPB*1KhK4h!o`Qpy065SrRhqco|CEeFGP=abSKmhl3iwxYlJQ;W<J!=r1|&J&Axa3qPDCvIQ>et9ueA0_tciJ*gN1;;`xr8'
    '17svnWJ0ExBppB8Ccvoy$qzro<0X|3;^aqe)<Rtr8E*vaJ4rTAO(Y~F3%IDQsbw>kh-OvbVH{Dy)Qsx~%KEv_mOW9ZcXhhg`Aj;U'
    '+##+BJBdslAWN(<Q<^3&I%q^Iv_qgoeDh~e6Fa{0c1uSi$UB$aTCvNf<^af7esL`b%pcmcy>c3oY>S9J7GWexsB*{%i_s93-MnNc'
    'V%g<^K;oo~SWW=2>+l7n{$=5k!xw3@AO%EI6!s8TjUu^oOAaKc5Ygl^$fPIZH`Jz%n?F_t<q2Mk4Dv|@z?p>(sFVD&Jj=t+a+2_}'
    'jV570<xdtRYzHog4A*W;Y^I<Yutm1#o9{oK!C7VHesN39j1^n?MD72TCSb<yvTwo^tkgwcwh0(%0CGpj+aU1@$Sipy_;;|T&4z+B'
    ')6VlJARxl_b9WuqwM8&HbsMd;rweLUI&5#WiJ-~mrm2}#99PR(2Uq4GNc|s(fJUydn}T5Wfot3-qFle(a+o&9(ae>5^-HI)WcNM%'
    'n<cNq7NI3m542P%aF<D8qSgK_TxkEuoGIWarK7ml2il`OP_)Ps^X57og^&-tR5={}WdyYep-9wQZ_Cm#83U`W$Qjdj;U1X?BQzpu'
    '_9k~kR=%9(Ce`Jat&gCN<cw&$L4tCxRT&(GlF<wZ`#t)5WVd8mlJZIuez49<q{NS;_I$vJud*Vdj*(3TJZah>Ns9!l7M_#LRQHl#'
    'k;0Uqmi#K&OuM9Y8MN=xu&5<76QEOeNHMHAtC|IF%C9tk2m2>rsU^s@E8*uRIWVc|2+96M-Vx2n0ynE0{Ln$@phHFAhE+$eBj}>s'
    'U99_d9onp1clYDba=tV@!`C$f1%{xz;CrSDHyCR{sm7KhvCRu{phb5rI1v3SBCe(OUmk<W9lmf!Oa21~4^nG|5opPoWJG_Vh>e?~'
    ')^WJg7kS=`7QE<ef?SS{I468X0cgO;fnS5|qe}3NBn`c`k~w1yXN`(&DI0o*Hk;fqzPy<W#_4*iR8=6veehrI4NU=q1saqKmlZ;Y'
    'aER$gz3bSp#f8Z>cDhzXM}mQ&%F6R6LFgft0fzupNfILNB@E10W`}tm5@}qV>xzvJZoR-m{4X@t`*t-^BpF$-M~`eIJ#|Mp6Cy1$'
    'Xei54jS2?@k1hoTKg`}`=MMUte%di{Sr*IJH3S%JQ7+J%mZDK}<0VLnt#C|;bU_sWXe&g4`a|3wL}HXg^5F`7INn_YMPYK~A`O}e'
    '@9qf!#%+460luLig1Zs}bnFoiOff6Tf{I+zECbOjO}<VnLC!EuIl5(2B+!_0MRerdpJV@Zyfo=^w{>mGY@WILTZv$&*CE4P@gKOc'
    '_Q(y<ndvIiUD*cp6cw2A?*<s_ER@kByT_bBsLe&?nR}a5nut2^>k4-0LlA!<GUdK2#fZ%z!%Z?%nv8E-Nt)n1Z0=woY(p(UMdP9z'
    'y;Mx#*n?5zE{~szgXc#^oor!yOsN6iL!Ih^^Cklpm5S*hi<qDO7<O_eeyufDdT5Xg<^P7xO35}h8Z1$!Em>#&1!bKUj?V)aKK|h)'
    '-Ls>i0<}Qv<865&k&~g=qaGVUhG?Rv0RQ9RSs#=sO)_F~<`_pnWUJ|1MRyq$ltP0R1XpeUNja@E9%&Kixnwss#9B?L*|s*#$1^7f'
    '3WznhQA4M09ek)a{3|PFoiCh$u%=;od-fR3)#IC6)=bR6CwJ9QJTi*Y2|NpzTDEV+Du_f)dn(KLq;(}{e1&Ugt&G=}AtCN+5Y#A;'
    'm%xg}Ac4KaRHTs7s!?*Fq8)iNq6nyAd>n<yMUBK{Fy~^4{S2|+v`#B{I4PjD`pK<NQ>Rym)t&lCOTmqTWpQ;>(j&6d{Si4s8A6~m'
    't>~1+;}6NEFHrU#OA0P$66DO_L$8QKrn$mLc+ZlnsJMY_pHxX-3a6rxY7#H{{q|DKV4DmjMw#;O%fYR&Q#AuyLt*TP-INmaE1z&X'
    'fdyE-38ox5cPkiY`<`xvf)*}E3xsHO<HRQ}aoXe`1mErjt>j<^)Txll3P7Rcw+s&xt>v%OQ3^nJ`BZ^kc2RgS$mj>{4W#vyrrfqU'
    '85RKu8D&VDOb;DGliPmetSi1=DT}gh_U6mkxFNZxBt+f3z=3|CwVWiJ??86xL!db|#RX*F(@3GhF5H!VZnM%e+1(Of=)!f@D(|_a'
    'hT`gc)gqV?fN-5=;mmi(vI7Yo<X`xzr>ZrlFz&LYU^sU9^6zWk6?6{C6qb@2J3R%_NoBe$`T;t681*i7j@KfeMG@U~CISzHRE|Dl'
    'dJ<DLx;*-2Ii{8wPZ~jCDybthu)uhWVh@Zo4=4(lAvActQdCdWW73zG%c96`XL21?xgjeMxXhYnuDh9IFI^#d96lik?uO=94<t?)'
    'AXAUv7tPv<8onE|9gRDY!6-$6?(SQl!CC=zw%jhCaTBLF>0E_vdydIKAM!tXiXAsbt3yVP>yUpWD;S}f3L{v-y0w!c=7C!D-Oo-!'
    '*9R1w$%K7hp8N!=0`cQHWz7W-sv)=%@QjMTxi%bwQMun8QU4qL-Bx^crcaOUZbVgOp>9LBT!!w7>@<b*WIqIXOc@8-qleP$lQ4{K'
    '_n*G)S7-|z<R%b`8o*X|z``g?uJ7~kBkK#CBWYVI{lquV5Xm~{VOIczf&qcU2jm0o*SB>js=_rgr2LkNyAQXP8@+agX&!8f?zu~k'
    'zOl3uXx*+v><WINd5h9%!5`EO!jVk>!|;#~_GN{G7!8uflVYeW3I#1oQ`m6JxY84)+{`-8(j%sk8An4CASJcl-OIKkl5I!hP3TIf'
    'Dwfu>;(=%g#=RjcyCe6C*Qk+usHFi~#k`EXwD^NZ=Q$xpb(Hb43n*&gS*$#9(q$1Y0|x-mNaA&BGLI9@iH*bA-$Ca@2Tw`x75zn_'
    'B13cfd?%Vv5=<68g-R7_f|<}5Yf8b>7@AOcnWmsT=eB}iCfd+E@ps2X*w*c62WY0uL`2*}>YmCl;|r+_c`;voL8$rgqVKK*)nI~3'
    's%&RX!erV+3Dy;>od{_nTARAmi>!ab^wA|$tDrSz<_asUnOu_Cc2F&D#;%<f(bNWY?w-e6Q{3r8%S@Af)fk>hEV<TZ*wL=WG%OlH'
    'C=&3#5CvwB4NRxa700~%E)#lC)(X^UOI1XNX<+J`9{sa%EIIUL1wp5BhdSsh+@9UX;Ihd!q3*?t0v{ev2Swb(5}~JZm#gf~SZ?&i'
    '6lroR#ugd=AUCS7(4;92V@6A~6eYT#GMy6?4P3%vFwRKd^{YopJKU)#XY#A!79BuKq=IGf`0AD-2XRu*y)Ew}!Hlc3DQ~k}BFJ{l'
    'z?Oegmb^mH#XX`vnfvvS+VT$&XM>x-hgxAKv+$(RG*w^a<{{7vmdH@FnbusvtGC&jL5}YV>vkSBr5PBOm}u{SN09d3nRO<IrXxr^'
    'Lsf(1N0}Pgj%%3H;E|2YzkFHxfIxnz1-VKqUKJ{n6_Rc%4p-#=OGz$#3?~g15x@B_qIk;<u}wfn#P#~yGN#PcM$<G|1}p27aVI4J'
    '*Rz>RR@d&{=^kPMqQ|e~*kb-G9CZn+9d2i)#ci`~ZNP&F)f2f}B2+vS7$eH>|LyUozkK`U+hI=r4YOhz_oayxTB(D<h-#}Va8I8P'
    '=w!%vV5YErK6%?pHCe&jwt-1lT%>BVHZ%`!dl9OHo);UCgn<16K|9=pSi#@m_#rwlQs4L_V1;BdPU^GI^-aW4QBJl|OA|EbW-5{G'
    'oI~VVW3sWEaA(2}G3r(@>UbbaFuC%>wz4~GjiVkTV$^kRs1nqHX<kf+r+0$VBA}l^-IfRP@F8MSllf`N%@0N<;E0%y1O1B6v~#eb'
    'u*<k{Ke9d>1Hh^=+r}ed_QOQm70^Qfa#uJ0i03*Y43fhJ5dp)2^{tY3jM(7}#+?*=Tc7o+4>c!^KSm=`Si+olj#M5RN;~6*^`Lhf'
    'uq~RVcc1lY`o&M;C+~rFw3y9GtbnjNxM7;SujairUJ&F6h$V2%d<TJn9O*i47i|q_F<Z0fUajM((t~t!V}oZ<$P(x!2|n<?#$z@#'
    'Oz5^DGFTRjQC9VrHAaw2=;gp;kRhWvqI|gkGVJq=so2n%Nm5G0jrf2cYwO#_Kyh8(NS2PZB@i#Ttz3Oi+imPcV3NF|+CV1&Sl-LV'
    'vp_SLKE_K-8@U4JN;pzTS>%OB!DaCwE<1QslAnvgbD__=BDQjvgf(lLV~~Rk0A%JynIL0Yxik6(jev!&b9oi3{levCnNdLD(FD-F'
    'PHp42b&`N%m`e<Q^;HG2D(=XL&ga?^Hnngl<3@QPz|xr&mEv@q2b9I+)F|%wp(ytYjfbXPv~P$9#CIav9^EoSju*}Y6<^5y0Kep)'
    '^H31(-yeL*+?`bHEZpe66NTPjBXOJIBDKNeGa1_FTIfXb*TgB?B*U|%@L0#0r<r16U9n*+Np5}Fa|kudR+D_PrcZ-;jh)jjbwm|4'
    'Sr>?JN6CAM8pyd;@w}8e5pREYRI*Jya#OGw#)3^uwG5c#)my*`EL<2^xT=)=-s+8ZW}<!dxlS{_3L;;gdOk1a0RPs5R6x*k3fxx1'
    '!79+w#=0R0#}(Wlf=Ozq*<JK}@;YpBjpP5jBc*=>_2dKp#2Y*?%;Uog{lE{|?<NV%=P<LSZR7w&Kkv`tU7xLwb*?(~jTG3rM#j4J'
    ')Hg_g6V0T%vaug)pk_8V2RR85(3BMe>u)-c06UMEt)kSy7GA(DLU@&(vFCsK5yab^)`tGFwj|sGZb@#nCDS8SDXrYkr6S6Miejhp'
    'xe+K^f}$)-40Wg?IFxfIydjker%WILrM$T#CwC;z)7#pSod*+J(?v%1orklHJk&{fPoOvBgD)rNjfR+z&jNLQE-8=(>=i2H&vmM3'
    '*A%%|>y$UikS=e&Xw#s$3V4XQa?+T#Dd8$>*o9tm!7`a%_eW5^8rsAZ4~%iTH-x<ZCf>*D_WGI{6g~1~>|`}j`4yS%$PCIyb&JL}'
    '(@R*751xch4MDpEpRFZq)bg@Xe42C=w857m1DgQ!nTEI?Kk@Np46V%xNE5KrjMscl*CUu^fiaYfZj|c4lC_<3ZpbufoslOg=19sd'
    '&?7<J=XtWumq^|cYS|DpA*k>_wkqpXWi%`H*Qm9|$xkD2gD?f0YmQ%jG+H6j)Xh6P4O0rf%W!+>V++R);mEW3xhCE$5d7PAb{h2l'
    '5XTj1lx52;k(>h$Kq|6ihooW06edSE`Nk;{dxSM~pRKQMf3`||WRuym#EHx>9hNQTaB0vHaBehRk&w$MH)Ux@a}1lbqn7F9jA40Z'
    'n%|mWE&Jr6kDV}vc~WRjP5DWCCbK?uK^wIXX#)#M_3jtMO+0eA=C)vJbKJ?;Nox_`nmg%A=3C-QCsDMY^)mZi&Mm{Rz^<tEq)}}%'
    '=<s$RiG8s?-DZSXXr4>`bd$LA0QFgRDwU#$rl22-b;c8pxU!VX%;3HNR;s#OR#uZ6jk6@h`DJTFb7ef)F6HTVdBC;hO?}yF>KI>s'
    '@pqr>j{V(a4j0*}=vW3TGw$?UL4qBg7sKVBYp1|QjDnu#;Io25AXxgRTQ5!KOWAog9B~*D(P2i*WJZAnLMMG^)3`^4;Z`PW83;TQ'
    'ZtU%);zETX$3>oMrVh^XAJ&0_ihz&Jm5c5bz&)$X2Yp$_vN@ZyY|73(ArdI`ad9i8EK3tWO1vs~r2RdCB0vgmXPGenGRBI|)XE(J'
    '5-}NlMjI(AD$Wa}8GNcGLS<S^0I==wP$n<2O3S9Xno90*Ku$|`e%WJ<RcYkdPfCaX*;GAHbQoO#wnXnBoivXvCO;FnGA{AC&x+jD'
    'lj;En5n;}=%#+r!A8?}vd5iA;GVn5L-kBOmm@G><pIpt4*45YRNOV6ECsE_c6q8+yl6?5qfIW$S+s@vci%KhcIH2L6k0L8l`VW+}'
    '(1K|tY`AQ6&x*^R_jO1H#py}dfLWAhPSd}%g6luvmpMp3iZs5aoblyG2D*`sOmc-lmHAjP(IrwOwwa-p(02PAWZ>jc);V?0#g08Z'
    'm$q{s%;7BidaJ9?wE~rs#lXX-d%{q7@J@D80ls;q-xxKkYo<LIQ+7T9XIHKY3kSEe<&F_@XY!ABW??YM!sO~nB21LI&=9LhCW}XU'
    'WESO}Mu5hi)YvrI{Q+Rapyy8B=>rIiOo2hIymVCK8s@Uo1eg#W3u*gh6jnL%@}#$WpQJN|@>@30<kL6gOI$R3EU{WJQdwr7>$3;B'
    '`ga5#$@huChzgjNyec-)c%e4^*ugs3zEu*?w>j0{aEat5T|#CaRk_;yK9eX~du4sLOh=@BVeIY*JD@pHins3g_zzq1GF-|k?3mnS'
    '9vfO_;8&l^Dg(QEv~4VG1>FuhPYO6!6l`R7=eL*hxpYUN?#Mgr44)mggf{5=_TDB%!CX;ghV&{eD>`MLD*}=;j%d*QN<PfO*w{!r'
    '(u`1V*BLc$cNtM{+sUAjEiY6-WH>mT=3waZ3;lMbY<2MEAEE{$?tn4FFG)M<dcR}Gus@eo2M81;_J4c)>8GC=DfVhr)Jk-Kp~EUj'
    'h+b`KUOQ2$Wml}GN$jcJ$MGzY+cKiG3aLI6J1WGb3?Rzv2?NrXiI$iNxK@+ZQ(y}(eyjxaq3w|{IYgBqdR`K<7}9S{z$TU_<6rGM'
    '9?cC6QbaWN=s-_fx&)lGzRQD!{u~qT)?j6-Qdz`*Cct5U0s@FL*Igx8;lAWYB@}Dd3G&#mG<53X<rHVIcCYPFC8$SGIf>HBwCnkh'
    '(E=>oeaxU33S@L&i3R3;SezaEYo$es=Kug4MfXj8oRpbF-rMG^xq?)MUS!2<Y(XFs8>=AFfs`JpH(OF6LMW&cE0j<zK2PmlA|e$3'
    'wpF{A**dAM?W5uWN<cvT<KtjwxSRtCj_4=|726F4LI<Qs)`_)E#T|$o2gH0q-DUQCXz=KQ0$VC;j69Kv{0ZrNIFzQrE>3G}^WcLC'
    'us%?`9(P1D!6JF2<&#II%W7uSxuzyN%EO!ht2~Hm*tKz!Iqc~PR+YQO7~sV|@ETOaV2{RNd!kKVY`2zOK88~Q3=2SWK|pxsZc(9#'
    ';!ac#RIl+<YKZX72VVT*RnR9KFO)CsxFJdknfd8n)0DN>u_tYj!+g1wE_~;!BS+}3E8@pqJixZF`P!+LvQ0*U8*@8B2y-f?a9P;P'
    'kg4qyI_=^xUr<hIVYy!EK8_e2t6<t2b(q*<x`gs<DtIwF#7ntGWf@sS$^0YF2>{}*iz>*1J&M<g+a+{LCaYs2N*NTwj?9}9MDif+'
    'NW<f$3v0_u9>N#dcez6Y`|4&$^|lLr(2WBDf9%IFu4)=8OB|^akIQBvO@aw)MrCZ;!M0@L1ClmG6%%?*Ok`~<*W7F6X#N(`bfEqK'
    'o`8!*%2jmSMUd1q?J9c~)HxC<nn){ns%212ZV#k#R8_wu@IQ<PUfSG2GYTUhSnw*7fT-J4S*o3Ojh^H%ic7Jx-EF4)0_IHX#Q%h9'
    'G2t-1;?Kikz?3}rub!!ChX$hPB!{8^$cY6j(IJDvO%qj+2y>kt-{ICOtbj4?*6;cSfeyL@2^9%SU^opZR#y#&*C_x>9#DFNWVi$l'
    'lm3JB<!srrLpdU^!@uJMTmXiJZe%BcgX0H*>I34&l~~Y2m@@~MgadOp&lB`3k2Hp7i-qd0>{a(bHUFyS2`WRW1d!)>aL|c_nF$-U'
    'PcKRDUQ2)z;Ac^IO7dY=2S^F5EDCT^+35l5kE<%kr8*4YMf;w(QQMUT0G^0_IZ%praKETjdEj5aEP*4?%N@07!;b@fr;K8rX_9=+'
    'jSd4;i7B=RRynNtalvl54Pk|Eh%UND;L6*gA&>x_w<-dggDfkdT&URqt&g$eoq6uD^%YSdbOzPMk_WU03gPBuTUWHZffs>Y<69z9'
    '^`p#;zm_7V#RbF$teUxl2#67N?tVQQjjF`PWoe08!b*W^ODbp3L<&y;@eRCKz_FGEM!+1fU%?C)aCxen%%cfxolPsw2TKDHwjR~3'
    'wkj{4OeIfLP-KC5<fEE2q-JJwbxkq1Sw@f5N6ucuVKbwn^#TZ6uwA;BCU_Fco#J*Q69{;~EI~}L8=_iZvH;=5LF!x;VwYL}u@u*u'
    'TF&j7A5L$h#D$tm+f>T3FAL&SVELFnJ#pbnh%VU}E|ZI7DzMl5^7Nh5K78j0b0j1`;mK7lfK9`<O(eb>C5uDYXr1JkTvav*gSA51'
    '6D0DBeG3EO6xsR&q`ng8QNX4``C`+Nmg?gXgPO%8z2cA9tae0?U|0;BJ@6up%A=LR9D%|C5pHm(EEtE$03?>1<-_-R4BW#GtQ5S7'
    'y)$7N$ef`*h9ZIXsc}6qr!te9;X&d~e>OT)6O!Ki?lBf(Ul<~AIP$~|V6vCH$F6HQ`YIi7f=IDvC)h(Y5rhk|Aw+gU9~vQK$pQeW'
    'x>ZiN&p;U21g%?N4&ec8YYxJ$k>R)1FfZ(H|MBPVf0G}^6Z}GNLrkr*3o@?Iuz^0R<L!9#%r-Ar0XCAiPk-wAFxpz-1>gKhHZkCw'
    'b{<TxQNT0)53Hy1eF{RZd~*+UhnuSz&OWld7*$oCf5B`<QwVQm1B!Sk=px5$b^<nqkSStN<BI(!1_1e(AL}vGP!5K=|EKlygwh*('
    'aZ^H%iNF?Km=^ftHbqn(9(ot(NV#l4ycn`De`#*NB&i=NUQ*juEK(c|pQh0?E)TAv%O_DCljEZF0fG$|81X^{BJ$WUQvx0Qgb%?n'
    '3j@}UJH^>g+={j)5a6Jxn^&mqn$wl5?3t1KuC$9kX|~aWeu);w$~7Pq(G|J>$+r`=8V~?ui*mGoh0y}_SgvfO=QLD*xPrpNw$6i5'
    '#j?@EDosbMmE!d$eiZ?*$g7?}pUp#_kIhE&>6>H-{!_}llk<SKqvb5uer4iYVf@=fX}f|hHph%9wxUh0O<G&z)GsGdqbevVNUw55'
    '^pHQOs!96E*Bv#8Sa(I9e_DGaSN=9J`0>+PVzB*V=O|XYCaO1~q`I!FBRU2^nLrtgoj}A^1d@>LgqOh+!FOwt38|8F!3bJspc-9u'
    'Rz3!Z-+>uyE2r(dbWDY0ZB-E8`Rm(He?2^ZV7{p=C$TqK)&-gV%x$QZWn_5VB9%!xdN>q7N!uS9iV_ZxFQ-Cc4to>53`;^Kf*B4d'
    '&DMwC!he9sL&FakRM-NifIgU{HmRY3Zr}37TossK54?(b=)-ZKg{BDJ$RUtK0py#Eu=b`YjPbD^yi{S6bMglt8Z4qm7;JJj8i|Y&'
    '$=~<DtI8+HU-uj8tbjDKr$EtC-Vy>>{K9g7Aft7M^$gLUAr}AKpE}cR$KzMw>OXu;(Z>|`m^iZuHYnWB)aP0~4mO9%w(9n#xO)6p'
    'KLMK?EDaQ6=sA2@yDuxaC-!N_r3ya}S;>h{pK*fOG2?YTMbX5z<~yn;zk%<lLY6~&vnWG!qS0J~Q=7{4{KvXf&Avt@`8|A9iL}6~'
    '_U3=^!8RXk@xiV|0+9jkjoxnDPivxZ<-*5{FOTB-dN*J17W@`I*rLJK;J5Hij-V6PXspHdT~OHh!QLeNL5J}F`RUI;JDcF&fBpEk'
    ')5Y>P(&d?1@IW}VH+ov{5(cK3J+P@WiimkeY8(*GFlXF&+*-W8DGENTw72PWG?hj|x%PNY?m*{L6deMPKBT63azSb~fX_nb;eBpV'
    'vTLftJ~p#JU_>2Gc<CLCws?jq_%x3vuhZMqmBP&XeDW`)@o9q{tg>?cG~!oTE58cr-TVN7a9f9!!9QUCFjq+x*gAZ=WA-}=PoLz|'
    'b#fY}zEJ8qN)QsGAjt&&i^!<!x3R8dXF?+!d8h0nizy#uMX!t3?g@M90V3tJlkDgnVV^n6f->q8Rd|nCy;S+Lucv8xmW6G_l_sv&'
    'PX3%EYr|PGevz8%>EVHeHy3Y+n1D&A>{3gdQ<D`PRpc1(x4pzwYj2OI8pr5<{mcR+EJ4IC+jX^!;uGLuc`}IeOs~sKKt*apLlMl+'
    'j4EZ4UQ7xZkTq@RkXgGINk>%k&JzRhidd3;s?z_YxJf4C8W`VSzVGQ6!DY*|{po<{^dmk{Q^M5t;ufQ@nW)ACIUZ&PllOhE!vNr$'
    'yGpaF%C=61?1q}#ytNf~drRDf=E>H)b%s_natdOcUg=JQVL&IF#~8F0pX&hG;ji3;sBxORmWwT~JqG&baVnCtRm4ONo^j9&m>P;~'
    'Tgh08mDGmBikJqJ2xXH>+nux|Ki8==sjri7<IZhoIl);4w++aQxq*;FZScmJs{prp<;VcCQP@QJ&>&1hK~0~pzGW&N7rqaNo#jvb'
    'BD=~?jyuX9ONyC@ADC?0tx>oH6*%R{-k=%o7e7bWRGi2Xxit^dP_+F+w;yFWLeFg3n}asEi7%d=?_1EAxvMRaeLTLm%0vTI7Vavk'
    '|GBnBsZoO5Y(+_}q;1j7uA5w?F$8q|*itB$tgTd(Iu#8~WItWhrWY>&`<!cFF6ZoNTG{G$AjHf^>0p1^N)^8_qmXkQtD-rd>%fj='
    'HYNUI`%ng)4;XFVGyl>@@Vy_lL>PIA53JnJ6oiP-W7_+!Oi@eb<m>KII>$79S^=b{cKrT>cKrL_F70S?CM2>kpa4_h)pz+*K(8W~'
    'a`YpjH5pKLya{7&^@#8n3w`-}sYjG;k2-l}(@65G6sNS^)Ih(EH$}jq4>_E+i1-0?&{mK5mYMzfdKE>k^ds%~O!pj#ABN06=DJX;'
    'NN3_&p&kR4&xr>2{L7^t(ZbyerJYb0)_OjcEH66Zx@eChZuN+(qS*)kWD@J>#{t5I#xO-X+7aKk)6HiJLKc1-YjiDYbR!7mtDOuH'
    '3EGXgR3oCC9CEl=B`B>Gte0X;o#D^43KB)`&;bCC513YQsY^be&&yc)l38VZ_oz!GXjp3)m%7BG0e^NVMUeo(eh@4`0ZWFtnU&?C'
    '{<h}9?8LIBbSq0#$&!yOE?ucLB{?G3G{#f#fKi$db;rh29%ewIOEccoc_3KWD<gnn2t|{n&#gvL-G`IFBEJL0DlBWwtIpb47bhb0'
    'g1@M6JA<X|Q9XVv^~oY?R{JZ0kbJM0MURSBVBumpl?~V`C=H~;{e90*P1<rMp6gbfDY0J^UeZyW_<oPvKQWl=pSJ4M>K5MCrf(`p'
    '_2$DQKM&$-BW>l$(<RAZ6D7QyHapDqSd`*=^Qp_EaRnXPi8pW%S%Wb(bin`kz_bAuCzxPNm-5sQHN4MrWRs8F4D+KsS^Z98o~=4j'
    'gFVE?>J>2;P3Dvt>5PP#<Z)=sGPmaB$hwA{JJ<IYM=b#hD+tt<Z=D#w7aO6Z?a;$^i@?BO>zU*_HN!~Xnt&T%ff`Gpk|(3-Tyukf'
    'E6-#-c(L39-;!DvH37zh^?i#B6o7Gdx{q*ev<ePsyh*Bo`#jquOX%0+b*PjJ5ti|z0Wj3JeudyQYk)g8nV{=izfz7ho;_sV<5sU3'
    'sc{oUfNhMWTqQxPOZZ-e7v8ERxoRo1&aGa#ha#ULm^zMIyYd{b{n>=&{n*1zjl$RrlaP*fm4cD1un{E6i=VZ=s4{ep^XM`9(UlP2'
    '`jvx_0Gi333rzNZ|MG)EHJxm0P~lmcbt_C56)L=z{9p_7u{A2c4bD$}+no~p<w>Fw{9Ct5wJ({eOtsuv7rl#As(zk%*ER(uph^u4'
    'YU%Q6t5n|BW{sbR?oP&77b@#QneniXxi1v6KUflsS2~q{gDuQI1(99P^~{qGE!404x9C2dO61mUb4MUhL(3{T;G}Ct^YFn+>To`u'
    '4S2lA6YU-H)~)_<UEi0{`EB0>^<?~mcmwzgcC@Pxyg2c^kG*r~7i-g#7BPmt6)YP34Kw(o?{q0xbow*%;Yldc(y)3;f5h#Pf;5yH'
    'V>wd%ePmZCzO<!hpKL_;{J}AOvZE9J9b471=^MP<s*1;N`{rPM<I~F&e%tq|mG^2}-D<tfPV@CnBOUJ)@ELmGUTMQe=8i$HZ0^n>'
    '(>cz)JFD-hcr@;lY@ZnJaa*4tx*m_`q0Qh@$HeE1Nx5w)T~2z%)4NN>g%k7%S1re9)kz!L@}|%7*1EV;NxUHwWZ%!Z8GbN_c)v8c'
    'aJDhN@tP;4!>^%P8lO>8_|r9-U8#+#LzqH6cwjufS@6%5jr=@E5WsCNR2ENy41VVKpCTNS;>BnxuQ1})ySx#Qw>pm=QL5`NK#y~n'
    '(=ErfrZ`JMNWsr@uo0Tl;~fqUZu-m}R!;pM2E*rZ#r5V(hc**jx{D9^naGNbk-Njb$A>m}I--<))W1)Xdp+|q@s>DjO%vN{ntf|v'
    'J$LoJo;*&(!7wrX!11KES}v{W&!>lRq=fM-sZ%E8xRtPGc8#yln)tqt3#ub{>#kVy-gqX@vZlPYQy|BSrRsvM$)A$+NK6%Qa*ms;'
    'MJWDMy3)rMd9kk(Rz*{&$A~0NemY&>=N9E1LPwHFa_3_!qzFir&T2a-TpggbEu$wz_VM<o$R`x62Ty(^ve)G8BazLW_QP0skbd@c'
    'Hc4;)mgDA<#`x*Xwx3%kWN;(%0(N*z6TkoE`#-q0<{E+*1+=GC)a~xtBKwLyxAAgqGkn{`F9YMe6|o=+LdY!7)@-3jGbWquHE+*0'
    '+q5DOg0CpM>z-{!>sx$W+U41rFpU27Qry@Z>-+FX`4`zqbA8_ErP=3o>Mw33^G|90DIeE1F!CEO4a}ZOPY(xwt~-L~Ucav~^mci-'
    '7(J?7IR1KRU^5hpos#BMUdy}1%==Jkq2~CuAFGYB#(0=2*?2E6QYn{5it>hf^e<bR`NA#rOPb<m6R(wPU!c7s4TF67w+a>D=`Q7K'
    '_r<>gtGmfL5&f0$=|6m{Y`sgIO%(QKX<cxEQY)6@-|~Td9jn_Mx=DSYy6U*MT)r6B+${Rjb^odzTQ3hc+tCT!3ueu@IGAf+v3!Q>'
    'rP-3YV0+29qKeH^{||4Y@a`$u#2~^vdbSeA!d}wzQnc7b`E|RtW??k#PIEo#*0D@3WqFt1_ry%!$ki!td>UbeVL@}d{960^Xui2;'
    'te0PleH+%o1~D5&&m90X_iC3*i_7edd0bjN*TWyq6)VFgOGfyT-P<~Q60q*;RzH@b)XSgM;Zp_QKB$O#tP#f(493RuSy~!ZdT(l{'
    'y!qYgT>I0NDY^=oIxY@X{oJF9D2>t6atOcp*$U899+6x!4WV{!rmm{D&%MpijZV_8A7}R2*+cGGogVq;e&=Q+ZKQ+p;D<ns-u^6v'
    '%FyVFN*uqBvo+e9pL{vA8QDPXYg3%%a%l0Wx!n@SRcQle|A97+`EqE{LNg8DeWhIvEmqREQ<)w}(^H^l7Azj^zWFaUpGvZ}LNU|H'
    'm4rvQ_cjWn`d(4FSKPv#q^(p~N*kW>hOV8Lr<>N{h>dl*Dk7DimjtCj7Y>;bD)}Cc!u#B^JqACsHQEU<Rr<?+{i9P`YZ#0++8Y7H'
    'U6m$U1gTvvZEbd-A9_tkk0wTllAF42!BsT1&iwKi>2_$Vt?$qshgT5F)TyX6<V(C9+FEDPT}}0kemS%i%<heUDSl|@x)4gd_(n8z'
    '%bm3ka`L~VOOYnpUUBPDLSmyf)>yx^#~%8r*@yFTW9<{rjPp(FM~RZ%yp!Er+RAOZT65C)YL@Z+?c8KmzlSI^qXX(Y2|W8aLE^~h'
    'g2!()Xoj$zGd8NYDTqL*Zxu=gI1DOCDSz(CZ@O_kAz}o+!>S$E9w$k(!)sSc6++7ctEyT6btbpP)}bhGAjuFX?%Td+X+CfH@>|Ix'
    '3S?u+$7{u+*5Z>BJiiy#d}1b)CU8HYwhWLL-zc)UT-ItEY4EVvo2y)xoa$iq%7>8^jQyFs{o`m+G?F2@wr@SkyeJE=(kZ<K!4rrK'
    '!|klrh8PIP6^nN)Y=Xx2v`eqP?+F@fBX`YkTGVIEY<qyG`GcjXZCF!qt1axsZPDk>zID>8l?uGN()rvctC}04DUya;$yG))!{L=7'
    'yta)65b(LL5IFpKyj@iL&&Kc#KA-v0r5arFtL(GdiwlJ26u>+TY}rY4@FAFCXGqlVpR8&s1gdBC3-PRK<~n7q)8`(SDXgM#O|I%M'
    '!U>Z7BHoUv?WR6_Ta_Fy$F#QYLZF$YbR5$-=~2h&6>dK}CKS-PcHB~wpux^S2g-4|rM3XsrKKsGaokd@%v#dvV_aJ-lqxwLN*|Y7'
    'Y7c6@o_2^ZjnbuPT!xAmyp+o=H3<#xMsL#Xmf8h+*C9&HuaunB9lCMB_dmBpt|R(UFGVKp)+!1{;Wb@esVzp(A580%e0il6fL`6a'
    'IVZ;+usPoPO^<#!_7O;u6A&bl71X8~(TePZ#f}A>`Qfs5d8IZ6&>Zj$@qy?AeB%Qr3U0{@(3EvZ`jkYk{IngF`9zm{{1cDw89?%i'
    'TCe1-qdr`OH{u9Fa1wXXk4!O81M|0a<phTtKb7e&Jvj$cB*?T#n@bns3gQ4*q@f#vqhR*Qflgo;$e>Y`7<27rzWf91<43j4$sR=f'
    '<`spNp_UY>*ifHPwyW+QAL(G99<7tMBCn~V6hx;nSp$lk_6j2wkKZ=LZ_5=yjf_FQE(VHcWKc~Yq2q%w9PIX}2q`8D+XAQ|4^W2_'
    '>`QSxFSdqzI-8$|I+?0^y3y`|PB1?)7In}$>{a?)&SvyDNkN!IvzpSH)uwfw$Mt$DaGG8Yj%5iY+BXb+wHy>TCaol&v?KhV;6Ni#'
    'tZd;=oDgqTURQt;=$nmwd0P)|AJq%d$UWt|^7v2~K5J#LEpIeUXvYnUYiZXgPehW=9%AG1cu%ZBdagcXYkuJ=B{}KtK1GkK`yQ>n'
    'o3plPzT71IHwMXOHpfT8pb@PJ2xZqxs59m5)yeJ+6}q=oxh7{2I5)dQG>)V?kjC}=%_piS|8_pYnZ4WRs~xmC7KyD7fX>OslS$e;'
    '{iNS80I}aVuk?qnRM1XdB$U3tr!Rnq4L*GLaB8`(@duPj;hwniXmfhQ8x%%mmy>_~^>0`3t|1LD5wK}!Z5uq8z*6vJ&H{4Kt%U0R'
    'q7-GT0LXT@8K@o)C!98bO9LCQy@T@dVJn+G0EC3CTLIurS1KuXc6NZ7+R-y$cuJXn8zWDY;R3v6$~+wGD#uaij$#}DxNZqFn|Xz3'
    'FlKXr0-3PoP_+v(wcc_l4w02OnyI05&wgX7cQKj>d5;aJA`2*nQR+bjnr~?qEN*3nuxVidejV3`hXT^~0kl(D2wQHCKmGT={11Qr'
    '@9h77Ldhi<p9N=eMrQL#*ux;1y9U<ygTs!aBH@H?B%^^kCm-%~vq=Ic3(eSDxoksQl@gc-DmCK)E=ds^%E=HYj;#_Fc<ejC=mOT;'
    '=HDE=yeRnp;St6Rvo2`++FAi^yiqKblDCEa_@_rN1%QsLcEdIU+Up_X53Fl-yZ#@32Zt=rbfK>8A&|;~)^C+vXpfO1+$D&<Q(zpW'
    'o`aJkQTtdTY|0A*Z5x%GaYdzb16eqoq3>URIcEsM96{Am!g0@m5Gv(?djE1YLI)ha4o8<e`pHZ-{MsP}T^}N#ZZA%Gy1ohs0P>&t'
    'WRm$MUZ9raTjL5EltLmhC$JG$F+uWg<<4)F1Dev0c*P4fjA?sBuSJm6(D&t>Zp%)%ZLd+z5yc(!AdbIrPLu#Q_?*eYa{K!%qT&Md'
    'JLe2?f)TeXLM!U@Wi0F?J8IrTt*j$n<;{-AP29q%_;}?R&!R4f<2tj6Pnzny&MM{{%~xM$cit#m{=7yXW%JFzbuVl7R_VK4Gso)~'
    'lV}4^<9yUCal%ul+oemaSxKu(f^)-ra3+d})T)1)^l3VjKU@GUaV-}dmLvDtN@;Q1ax5E&e)sBaeF6v%{Cnbz;xf-(uId8ZyzAz&'
    'XgF<PnFvJw9I=1_HSNqcwnWJ4jevRfowz1*YBJ|>ckpTBrYfVHeuwDw_Gmt;k97z^E{pzw^Y}6t!OS|a2j{}e@zXZr77pAssa!$<'
    'zwU$M@5RNz*LS)eAKv+<bv7bA!7)JC(m>h^?X7JvHDAPmXryGe2mYGb6eS-AAcByA&5LUGs&ud8bLLi&vm*&xuHw1Qa!T-QR`lu('
    'EngZn&kFu~TZ9tK5p*&#4_cw&38Fn{8qsz@ihUK)2O=wWrE_e~C?r16ad0%p9WyfHg~j|Ku-xV(P0$S@o1X2NEd!#N5wfm?<*k55'
    '6;jxxwmW%=*9kyr@G-U#(497(6hxzSVbj!XuBL~`bj;B7GuSBbJEb%Gg-`FK;77lwva(`|KFwy@6N4Ox%D*&5FYkljHM3emg<;d&'
    '4u}p*SnStamw-UpfzjE<!MFz+SxvLvaz6-aTe6n$LpBH*an|a9D!*l`Y82l)sxnVj7*!FrNibz-)i+!9rRFY~z{5^jaIdvPeET5e'
    'U)gQLY``3v@8$pc^RJhAh>|NX|Mb-yRV6)*y5K6iLw2y_q=Ij8x{EB91yX$pzGbZ=x{k1AUxcqt2?XX9QV58HO4cT72uLiD+Bn(}'
    '-)7^S4mQqwY*2RezzzoQRo@w1TqRUpM#Qt)@Qp0-__)*4Q0K;a9?%aWND?3O#wp{MJSzC<v2#&h9yDb^%vLJkNhmyn5*Tp_bBSUs'
    '3@=lr+w)zd5!6Gdg9|InOcCu8+7B!O21}Vr4L54The?Mews|5us3jouK@oCnXgtN#1gL_R!RAP9nf*rcduo&}XqVz{nNrskiGs^j'
    'I@wNKzZ5#Xcl^G(1+wNa4$Wb!ClkxtFP9^&AasK-GXW>ACkG3H6D_xZ>28$AMpHM<9}I&=11$E?Zqp!lK3Q1?gU2RA;g?ZZe4F&N'
    'As%L!-39{UY?FQuu&9@b(5ceRlATxUD%h^%7fRzY9dOvR{Y#!FT7Wq!VQ7UM3oI95%4%)%b`HR@g`5DOY4bP+=tv*PJ9TV}tSf-*'
    '=5_|q%)O8XC!abTe2V&Kj4rgX;|3H3rCddh9PiY4XE=_8X$#TL__mVOVO|P%5KXaTcVzXawz<D84>QfUQUJ+b#s|jMow!H%laC!1'
    'm}qwgin?%MOKeuB&!czwNKOk>IK)!Dy|K&~P|465ZQ}?32JIC1oH?Ip0(L@*+6kdvSl~{e1;$w06;9~d3vq^JH!Bqhf0HeI$WVHD'
    'N3SSp=njLdAJdz-${0Sx&n5G+7Qz{`m#nsJ87ZrSGZ?RhPgx7!%VCqI@MpspFytQK5vQBY#X)x7POcQ9cFELq0=apJ%jyRg3qbA3'
    '?TXCu3g!O!2YeB$sWWqFV|6X_h{qsaq<Wa#6T&^V%<Qf`#f9`O_~h|i+;kgcWCn6$Ue~_ee*hUxL74OG?b1guO34_I(+?OU)7?oB'
    'rcv^^Fh&KD^G=QdlnNh^5vsv~O6Ij{zM#H*)<zh5L7Ki_P+!}etjZKutm3z-TB?rE_1l8`Ci^AKM#&e@2-bRN{F6g>7tkn>1LTP^'
    '^?XB(N@RS1&B(Y0jx@3>bv>L(#xXCjZ$ifo;1Q$ee0<R`CXBxxvlGF(h6tjA_*FLxnjsqs)R|agg*51>+Iu`v7Mc#|2v6&edZCV*'
    'Vq(p$Fd-GtvFd@Bx=U-GgK$OfxB9j>Oz*@nrVD-qm4d;7F*08jHnw1iPzRd`l_|XK`%T<0DZ2PE_9nvi)jrY=-XZW-hK79?+K^BU'
    '`^#+Pae<Ibw1egvZ%_w!Lg5LWq->z*_nR)@Z%_;ubdK#5+dMn@Jn`biw}K>r#MZtI1N#T?FO%|otv4`f^dywTm`+E#?n(sI6Nc1{'
    'PlAdrK5fb?E&T5!#nCgbKJN>nR8a7yt(Zz47evYSfH_QAbK3LB7X%p66-e6Q%~n76;lzg4;q+SH<aoE938wM2)RY!>lzyVKZq65E'
    '$qshVA|yMum)=BhC{VuGR~$0BmKkMxJjwk+E!l+h?x!}pYvc#YrX1%f+%>FrJW%{(FVovzYBh&)1Y;!$w3;ml_}Ri@5m|#w6xkT$'
    'pr0TF6N8<?6#oEXvbn92nMu9oH;4(PncSo^9nJbgMGl9<rbipaG>GII4y5q7eGyE)fF*KVI3~!><Z7TW+4OxKr7@#~S*)SoIHpm9'
    'hBHRV>B2FY7z^1qk&ct&m=JY^bXiGh2L_poHMA-r$CfW36POc->#<9gaDtUj&=VrxVWSc_!Ldx-rLpc^nr<*t8(r|bo^rdeOg79J'
    'e1GLRFEA6>aJx=_Z{x-?P599Kiz$~&KfB&ZTkqtJWr|s+Bzt~mp9M%TaNo8I(PZ<Ub!&5m`*7v`@BeWrZ+HTT>*NK^OHMsC5CZ~C'
    ';2?krmR$f^X=i)`?2Sn~-jPWtWn;M{-sHpMz{@s%`||M60X*^=^Aw#14huIs)JXUS8cjNdm@CA<h5^X6#Y13%?|?Whf&BWYJYG)#'
    ')_G%T8&QB4ZZ@zX*K=bm&U3MB25(L(Baq@8;Kp>pVj)xT*{S8kEj)=Ex7uK!gj3zqf!$z=3dcX(LUIH%f9~m6ffP`}$#m@Dz{h~}'
    'nDJn4@Jz~7m#ro64gFd{y|kt|4z1EjpW-`$-lhYh!B>O!$-MBUcx>f}uX2to;%P2(z5u5|<Re$2l;CFk>rPhjeM&a8>2yAIjZ~k8'
    'E24_8vW4ziBwsyW`qZ}&=Co<{;*r6pK<!9iTrFxtxe4k6=-QFquwyOh%=HZGeQi{<yTx_GRbCND<<H`$6M{|j3;1%%7<n^Zw%3DF'
    ';mK_p*zsEUoDO;%Cop>*>*Jbm`Srj@((z_|UbP+Z$Sfwj^JX5x_RTtV*_q=7u?(Fjpqz3A47N(z)Rpf1+t+`)d<hJA(ka3!Al$UQ'
    '2^J-z=;tJHkk(~WFHqZDo<w$=D9^>klDs<;Qmv6B_1$-pupnqeG}+$JwCmUuQy4|;Y;+#$xUeH=HJI4a#yq-q6%HbPY<44i1DQf-'
    'C&0x^b%cs}v3{XT__x#vh_SyFb%1zx*Z@Ee&LFnDiAo~})VaHK5Wh5&48VET>Xj3V#;ndkEyO~^c~>r&XG%1zy+P@AKQ0Af_Q3w('
    'T?4d!dllToRgY3yPU&(h0F{Ip6pQm9$ggO$)!V{&=sB{UqrV&zC9F3x=&<cbHgy*hc0fc?bk>pXio0nHpPiT+^qo5-Hx2ZXKtG(S'
    'eY2Ax*kEBfbW#3AXYoE>@+secGJ^@gKghOk9|P_@L<i8<O8$Ra13{ix`?s2GzRC+Lzy%q1_6D*aYc4^iIrWnn`&D3AX<^5{3a-`q'
    '#_WYoj6*^G@dDhKz)8s9qtkPUy<DyaIyqyN6Ov-5>ft%j72Mp(HeOeiVXugW*mU&nC7$v%V8ogZst8ry&ISm^SO<s9Za)47xSgB%'
    'Gi>YD?;8K-B*mxQx%$+8-zyo0Phe}c)yw1H)nVv>Lff$m&9^ElakW!dO;%Xf>$B+ycU!JYPcfMuIqY7SyWoe|&NT5}u#4i4eup+E'
    '7;Dt&cFi#MJ3JEf?^fi&1FI$pF+2`|%@L0G0r?0}0`Fv^);WiK@v{o=;sO2ec}jFi(L+h;azG-uKH>C{X1!cKNa&eme3^*J<Arpv'
    '*V->SVn6ffz)B2=tu)x`u2Kq`7HH*4nIcEZJrNX&<JR@nPl#{YOayFh+w!3}retkn6Yp$7R%Va=c1GX~VpIc7==mGDBj}o)=mEz|'
    'h2{zP1I7JtfX>I(Cfx+P<Ww9=JAL1$t_S@ENXmHNMyAXG2*#c@2e4e#6)+T7ucs?J(%b7RdXorgy}M5;-vs;=5h+zLO4UA2PKmn<'
    '4FkCivKlq6DvJRqE<_J#0qZ8_k(l?rZbw{%pdU*eR<=;AvxTeBg7y_zSpv)b+ld`wOg(UYOcGq=0XAw-TS2)3Ix-HviBpI2PYJ9h'
    'SI^xOHYK;0<ua}_({yHrcSZPIdMTQ4aJ^KU`J`?XeEj|G_EFFe6Bb5h!=!$g>uP#9Zr&Z23>5qXn|^cvnlMEyFm6F*v+lbTn5}?U'
    '6rjcO4j^WlzGR6$96K091Pv?((bM6e<F{+&*UdG*ARe^Iimo%}rq}_da#x=DnFh_1tli<530ug-d!^){@Z^6<IqPntIL5t-qq(|N'
    '$`@25vwsvGh8;`C(%`sIL>(Y5)N48LQ-BGEx?Qzhg&4(}SI5a~MfTP3ZeOQX0{%qv9v(uNWS}$kC?|czb|qf%Ao?h9lCmA=j`>A4'
    'c#rdS+^WhDahZbgeOuQzzpC*`S{o;@*HL4G3I&zij<r}BkOQn%bNKrzEke!NsVn5RBdWySfD`O^XZC^Sj%~ay;+)A^k;VS<<{(oT'
    'Inx1!?~bYzL@+MoB0)9-hBUfbT<I(#deAh&%RAgh1D?Ngb4_IaPaF^O*Cs9cM1EgWWajSRfPJI8^MEHV#s5%`-sGx5ZgbscXctLN'
    '<bYDT8Rjsbsq=WCxq7)1TVRu)zWCl{qXPJ9SI(D%2lfGYGC<YMc}Q|N`dl>Kxs(&3Beg{Jx*R`q86=X`zZ(xOHmX?<phxLM;Rf0c'
    'qs37<d^m|Mp&mKL8$W~g90B#9nF+j2h?CG#{*i|G&fP7f_9fB04JWp7Ttie7#R-|iboty$Z`vimM3YxA_-D^1fu+5zTFB&N=%_&k'
    'W8R^d<2^@~*qZVMzps-%=+I}h>ujft5YN;n{V}92FN!(-r49ZeJO8J6am=(qG65m#aUp(kW|uMDLtg*;-l*87Xa<#+4x@Zye1`I^'
    'Uv4L)5m;j&$lFPM=@*C{k)kP__!^~aD4o9us`~lxEet<5OT^WD3IY>Zp3<UhkYE)9ds(ziYkor1z60Up(oHw3DoO~#0EObnivPTO'
    '$t}$1e6y;egK9&AE54Z*2*Yf=NVN$Z($-JtsbiN+Y)9I>_&o{N*8wTSOW9hh6ZMLLh`0T7f7Lv=&V=e#W+*yA%PGDdN_AeMgUS<c'
    'yvEv*gwVSjyS<5Xw@`hu^~y6Z(la@5otAy7{`kYM0Q;6#XrmtBF{Z8{VZb+nkhpJ}5o2bGiKNP>(<!p~6HM!T)633x;Ti*lD|plI'
    'bXN@`vMFQP?L$U(WMV5GuQXL6)00<ut@`>Pix}m>7i{ARw_Qw-Ad~A`@9}~4wA&wLXE1~?PVn#vI3rm%s`q%7KU!E7h?jX~ykZy4'
    'iD`p=ywTZ|G(abUuzgY7I)KHOZ=zSNIITuRN&+<5W9vo^HAF3bzMN1NcA=`@Yln(9C#r3+oVB|mm2&K`n)Ae?_qHN%eEN|^w=ht<'
    'bv)35?ZQVLuX_2QIVqKsQQ3HN-GHeZ`^(gASZs6(F|c#b6JFhMK~eV-VXfj#5+G4Gwn4gu?Kl3urVTAeIS*#>T^TPBjL|0p@_!un'
    'v!gjKoC&@e<)~54X<G;CFwH%C&`^gDbstNBR1gjc@%A~<dK3%XUgHXsjusgz9ZpRyO$tCl1bH#ugx*d`A#iH}?Qr=<ZHnF3WKH*c'
    '9h@BHNa>melN~3L`UFF4!=Y(OHjO)(=3456fLr%qMV??7MS22x;&@e?q)@exM1Qx#euLo^M8o(-Xd*e8O@3h$FUrS~lfg{c%2wnd'
    'iu`|V(m6BNrJ7O^wS`L-pmI8?9HvVaxtM6?0V(13le?G^-L4|Zk%E&0ek}Q)z)^%((d2z!XDL*zfw?0GSr3Q?FfCZ)P+~lRT<o2m'
    'd=vFpZWf$^ZbJ}Bb&D7=FWXr&ZHwP@)QBF}5^}KuFUvQh&YKCRCWV!9cHe#pe^SD6&lVMkFU4gpwiH*1IUg<@0wMJiGkW6z>*DMq'
    'y+tA5RHbXgW~(=jtDxv`t7{#RAq|#m+ZBf(RUx;^c-<+l&^$@L>9&_LX|qTng9LUZy?*`AqqIT(jwl8Jy|w{j-<TEf^});-#oeg1'
    'yGml2@g`SK`qEegu{ntr2RP;XRiTK|=mvw!Of__>X~{dnMu*SaL`0j2i6l{rJAZ~47g~7KK=l7n<*-~?gfmEan(6^nHk2IM?x+vY'
    'JlK{_TdCtRDOO#t=6)R13XDe~t{Fg%a7faT5jELv3w7t@k~ztE)SUhNeU5*;euf4=7-;l$N1tcYHglwQyjA4%n=^1HdyfWoCwo^R'
    'WfKfU5R6toK-BN(yb8c89KldhJe|(r1))#XqxTlQ?bDkKV5H@9U7Y257Lf15C~m0GIV`cmrWGAlf#|QGInT$t4SjT$F`<zaaPYGE'
    'KYh^(Ak28}c_UodQ0W1T_}m)b>=R(=K<RJGDb5h;{A#O#@Lek?iaiuNDOa_WNGJ=wu`US}V3D-*IaILsZM+Mb)glQCP$*_<Z0!&H'
    '`h>lJ$Wl;<{=-+&_Aq*s;$5$)9U^x4D+7H)onX%~FBnd%#K8d-4X9W?NOEYagveI|LJd!QT;*od3m?J`M`3y0VZO6tYjWPqvR2T)'
    'p<aBmc2?<4+vjM{e5?X{65vq+C_8(vw*#FB(FLWZZD3RfUg2q^lFgHEaL>(2T5vRuqM|S0010S<cf5H9Frz7eT(*L8vy@pbHd>Uq'
    'b3t%GyB^EC9if4J!_6wf1~zGjl@{A(_Axm171eB~?9`M5OeA2VQ{o0V7DGc(`~Y8Q10tZP&$83g!U`a6Hlc09302$d+9M$-gWeJd'
    'yc{+HeiP*i#GTUtISq0u`b|%<G}sxN`SAPjscnCt`RZXMU>kw}QZy8Nl~4}WzE2ONk1~KfY8P{@-IsCsWL`nvWTcF8O<;jBf`pp4'
    'e5ij&HmD6y*9Q1p!>%Z<D-RY1B&8I7QH$e5T-oexb(~qEJJyh0berqtt)6|;PbB=&v>E%~c3wV(uE}=jmTM_3xklT<hQP8$^bAc0'
    'P*bo&ne#^c10c^6!KS7zpqxdIOUU(24@3#r*&7^<Sb}np$ve|G&}Oq4xM^Rndt18@F>UNL`1W{vmCl+}M5Q4l++cmj0lWggYMUC)'
    'B<gi+N)P=lD&~4df)NQ78&p1ZKeqeOlpvyKOhH#=1=Q!+^`>dG>)n9aDmSvbSVSf@p7J0FNhpVbEoy34MFgLVV569yW@5$O*^W?O'
    '4L3co-H;*010>X}9zt9yu~K#ZOpuqh3FpD&0OalrZYMe|t{Iv|5a?20?P`q#FcCHzIvLnGTl5Ei`Sq`t1w!9~9hB@u3um#S1m>K9'
    'uba(ITRBkzxHWLAP(@j<?kfackx4&p^8&LxY?2=6rEP!MQ3BMbKus|+vP}!L_Nz#F*w@F%4l5t3NVJJELl!X;Z-FJV!Y-89tjo;a'
    'akpBt|3=BL{bi;)HfvB~a~tO-kU@dwpba!6BVYZ`Mg+G!I@4<0UY^k!s{l+o>|$Fz&@@LSVdx}A&R2m(VryA&mATT?T{}X(aHaAo'
    '$v|0^J-@kSQ>c3B{G<XbH0MrUlkmBwk&oz&P1NR~q6_K|B^ij^*;#5EuBC;9Vzga2U*k~zRJ2DzJy>6<yLAn*1fVC@30FU<5Yk<N'
    '(b?O+qBBEq{z;vQm&r=>Yj3Y?&d}8_rLY0886q`$<a?E2fHfSTZ)CE{Rfu6H4k{b%94`#2>77cN=2+>pk5i11kllT-OVUNr(BPAU'
    'jH;n%H(e|ZunRx|(80)m$nZct#K~DSsjGqDMB6M}2dvxaZp7EI@7dmNoU-kVLQil*p<9!}WWLR@R*k#d6x?BQVCf0^L8aG0nU>6U'
    'Sr7A%<J}|$he3_m%%bTrUajR!cBA7Os_2XJu*$9ijoPXx{oE(WY{<wHgobigg~lLhLZVj6TVsNp7NVh{qTJ7y)(h$SWa||dhlj6M'
    'Tk93qdaV|46IR6JKzpcnf0SZV%?2{%_OR^$0Bi<^hO1Bm2!Gbbvp{vLx@jVbM;`E2_g0rUBySJ4V|U2X+9{#<PBG_n-mMB}L7)#y'
    'pVds~-Op?z)ssyVy!NBPkP!eX&!in#V()H*)7Gt%t7&3CZsj$x<2hh@B3p4WehMKEfaJV9<zz+<g|`h=CDk8$YuFyf1EO~Bc*+cC'
    'gT)M|b7hCuaI^Ue7_J#?xn3>(I!t)&?HaEbHoDfLDRsIU^^FKZFUr%_mOK3O|8&*IBkmS5^TWtO4zaZ<!^%3lDz52IZ;^zQ<cO3$'
    'V9!0Zgiv{=pcv5H3JTQ)2!@y-M!2Y`EW%4Tz~|VCY7MJxM%fge{6ho<LmJ+{Vf_UB@FZ2MX1C?5vJl2tW#tS20w|2!;43zKjwEeD'
    'oI@ZX2VWr`47|Gi#RN<#cab5P@9Vg7wUNevTQPoH6T7|Z+8wPP4)O}e29X*5ZS)mRRQ|aD=Y$3a2wdL99PAf_=2@Je5W``P9W4I$'
    'd$ssOL1JVxlN4C_^kYe4c+LEeHHk$JAq@zYX_M4_1Ls`<R@*e}u5SKNb@r3BK1q|-_5G=vA+h^ITQPKy0~#;Fq?_qvcM$=-nvAF%'
    'mEg;coYiJ*=dD;;G2l1FG%xK4NeFXwIyLdfnkCb8Yq@R9^}BR&!1r~|9Tyqs3Je-n+pQJ-e-0Fvl;RzoT46?@PHh|on(R8+4EG$<'
    '&4XuJOv6Haw;fWWh$=rY3zh7_gZI_F3f!Qa4rf_d{3J7zzAeV!JH&H<5DGX{u#R)N^1q#&E2wBxyzKj$Li5cug)^=+nRm9Y=7JZN'
    '+XZvd9kUzO%ufhru5({i>Ka-g5v09NhgIxTnC*X4<$Fa<PFF4%jl3Q}(|KG;wV;tT`GxM(IKU}1>+Z9$Y8l;x<^`M@=uaHPduSqc'
    'w!@DHV+KK&)zNYta<Ja>7^YszrBvY{+R6u%<{iq}hJKbNnR)=5d%u(_3T?oOH@ntwbs$5nEy^55taoEa47I^V)dV+H7b<z&?G|Rz'
    'j+=rOlu97;YY>>`1Jr!-#|c!!=!1%lil&I;Fz2DMD(KoC8I2=^j`SpQXKOUKf<23DBjLAW07~{Mn*@{fU=M9HbTkwGfToa$sS<QE'
    '#2+YWC!&B|8`7)8KiB?h;C<EOgG}ydYRm<5v1aigKyZ*z_)%lqA~@S{=?-U}Ip#icwlCetC~Vd?mK?S@IJ{1z+?}Yq#bUN{yh$dV'
    '(v);DUo2mQ?$ih-r$9%DaAVD<uP#MQn|?IV2KdS^b5+~C>7O#htsP`Ad**|HLKpUd96Z}7An3RlZ2o#_=mLu#Y4HM_qvDJ2KQ`w$'
    'py)#GCid*LnVL5)Ix15con21~x&Wq+eiim+dMT@DJWL>>!?j1F7M;z+VN3YjPEOs8QS^0$HSvSk+r@Qbe~JDVOxaZ{NrwIOLr;{c'
    ')*k6}+O``tJr6^n7d_6L?5)WQ=UA?8oJq4CUI);O3!yTA1611k+gO)DLlR!5T}&KEJ)qX8xaYsBTG<|=|IiJEJb9Pr7y<FY_6@Ib'
    'dMV>c^c>|Xf-Ee|nBz(c(?fisK&f0^Ps$-|i=IM5{lqgxrL<4mlhIfX6gC4pAJRNrKI5pz!F;m%3E4Td4E-2;F;nuzoImtDQOU$S'
    'VPJg8B!%y|G%iXYTkXIn4hX$mx!}s|4eHyk%0~orD(ULeRri?T5dtKBdz@raC56x_X94}U$pzHZl|#nPS+@!2yoWo)izXPzFe9n3'
    'D<1jv?5h;ZO~?)~H~_wDVi%(j3KqKdRZRU9JnD$()aHC6j{zqqPO`0vt-TQeWy6)~;sZ*Ko(JrGMWCj`eP2xD%{zjNn<kXS=yAtQ'
    'Ks4dqa3aiCZa?%q*(8?U3He1zma4I*cF5VEm6>Z;7Z;l#8P^}2&;z6bX6Lbc64f=hpp&Mo&Myv^AvoC?H<ws4{ZKf2pie!0!ckI!'
    ';uHw3RF2n)mfj}D+7y~Y`P__JguJ@5a4vyqSAl{I;Kh415kiid{7`qGRU4vbO@()nNlX4jv@_*b`b8!{Q8RSb(G~5)SP@V_Va<{Y'
    'X>Amf*xI5XHov%(5JTiLpdH~&!_FK$$aJo5_mD5eC1vpwx)cc7|7kp~xSeci&K9#|Xzr<9gi8}GBPBBS)q)>u4Clce73V2kq})g?'
    'y4(mSiOZNc%2*Eh&65q{@$V^G#6ZiG;6d}U#d`D8e_d5VW*0@!y~98PT}8gsVW_)7h5M`MCV=e%s;H1IvNwplKwN^Ce{S=OEf(;@'
    ';O)JXm;fXv5B>p3i=`+z96Qb$&GuC3_@a>oZd|7mISl^`;F$PBFUL#u3q`+DICKO23;OZ#FG10c`7c^>xH{5jWBaK1JzE6f_p1r%'
    'BG8JQ8;HhwpjQ>vI|`42=TfCZF=X=5^NQl<`nK-C-gvxJkCzffIs%LZUSGeoXzIH^!X~*@S$%F7?5eX7Jt5?CUVX(1Th%?wLTY^W'
    '{hZMs^KLD6->f*$gAI_>ad2taK+4&x0I1U*$^8!1a!~9`)J5&`ZEu4-d;-Q*_d^z4^od{U)fakg;4pq@DaN|JeCTyy>ciK^#_P>n'
    'jfs5DuGzrQk7dah$u4S9IKUb&dL4AC$YKzPC#uV)CN@2_zrt}zpKxtq7a$-|FC9yIm<N?lKY(P8WFqL6D4wrEpQf8Ozv=hAPQ};B'
    'yTq!*Ck+mKQSnK|@&o}8&xv!0PYip(gARsw`&rdJr00ABX53F{4=<IYu9VZE_JeG?IC3?ONOi@t&eE{>DMISJhIbzng*whNpl|p?'
    'Srv)lhhuu`R5|q8k`KKdm-IQ*UJv}@7rrW0&U|7(=BVDL(yQjJ#vNRhyR4>Di{v9jIEJS*n%CSf4P4l0U7fC$By17!Z0tKz#_e`P'
    'N12)7l8Nqk7qAMEW#FpNxwc(uI8eV|eZ?knQvhH&<G{y~?ZJ%3$`SS)cV{t97gDm?Gm3|<y=U%Y%97aqqHyQMde^?`ZJZ#nS-p)k'
    '+YIF&5_bu0Nsw@5IJvEnPr#L;s+|{}3J)@fV)L?}xi@C&E{d)y9iC!bDFQRxPn`~BYt`DQQ?HYwCi*2B(>wuUJ}v;x9Q4+~5-zQ{'
    'u4w)0DMmworFz`3WRD>Qc(+E_CknO7(!l@u^@<n+VL5^l;AR^ge>4*J9388>vJEH#Gk?9Xl;8x6xxqm?kygEw0xa{@JE+7^k9ttg'
    'v9j$WUcxwBtVqX`Vc5GYf9@}XlDu-)oDeaP3j=X5yiqzq0b=|dKn1=j@gy`OAwTHQgf{}xNcF1NCBAN+pQwXA90f7_0%#zqD`Yb*'
    '@=dyIwghN|@FpFAl<lfPZfJY6i9GeVNZ^BSlC*oLV>VC;pf%*1)-L@E2~J9s;*ninslrl0?@g`YPcod33lU{L6@ta~vbqb47#<aq'
    '93B{-wR_fT6g}uszma%q@X_Rpfj*&b3e+W4t1J~v;w-hoJYiw%5Y&ARXK!~kj6zx%!7Ykm{@WS>5HlPE8e3m%_UIzTTQ;@C{8%%W'
    '**F+)en37mnhl+Bs&6A;Utux{iVB&F=?+An<l#cjXq21t9{tMEe2I-WF5df~=8q5^ZCp)08-omhLCtY@`HZNufKCHKtY0knO`TiG'
    'fOfPp1Cfk~RTiPfiw!?8Ld{#pjcwu{XNzghqFVg;WP1li5R#t?BEY_K4#Izd6#K-EIjvA1d>1_d?!)aq-039(`*ilaPG_H&axT&9'
    'u)Pw-=XEL?QD#}aoE9w2ybp<|7{Wn7W0Y**2RhGwb^S!&FslD?H7LWlQ2@eO>6d;6%g~T%)K41r*%MnR0L|Fr)%Vj@bFw!2b+sxp'
    '4`waq%GZ3JF#`C4SGiEGq8lw<l&@S^C7Y6(Cgagwb+`BFsF0!sybR@#@k3#$Y9mgkaQt~4J?JsK<fEsNnSc#6b6Z+ws{q?}QUM})'
    '_`Z%)MEFq>sh?afj8YS3I64sQ;ffJs7xk)Ll+)d|ECLkC!M2iC9#Z=Y*j?1Y+M84k%Y`Ikql6seVkfsp(1Av+l_+|a=r1w(KJCeg'
    'l$Po;Q%|ZZws3`wv9ppxz(T86KW+d`15gr>2iZ9_)`-o%Z5(kNS1g%2r|Z-ud6qw~B?m=Y+N=cYFn~=Cm3224IxLqYO@bt=)b{`Q'
    '{XhTE+kfcoKlJt=dixK({fFNELvR28L~nooa?!jaqk&Fz=-Ai8p}U9(IgpYa(m+W?2Md$Av9IIhO`9NgQN9frM&D2?bi3-PfGCHm'
    'Z#L4R04AG*E54}jX0(OnZ(K&l9eGbL$DIm02drSiS9=zKLc^$$KxQOaUWcjzINM;usHn3)Wj&~205~JP`iT@%9e5($)2%5}G(`l-'
    'j#|So2M1ekqIBXn!saB%Dpo;-+h!BhHxsWJ%4=LjFKB&LN!bH5Uptxvu$EkG(!h6pbs)-A3lN@H9qa6bsepfJobYM_|K=lsw=7>B'
    '&0a(OtH9oMF^u<tFAwj3^|@BgHZK;gq;_Z+5`mX#QF3Wqpx3f-?RJ($d#!>tf1>YzB4caiE0`AXI_eEdwWYBI>RKR4$q5d2n}`^k'
    'kH<G4utwm24TKzCeZ$zO0aJ~p*Y)Zg)_k)x-z<`D4K>3LNS1g)hp4JzX9^T7=c^n7^x`I%F^wYr^>T-V8Zhh`bbvawlRzNz`+x}F'
    '9ZWKB3VPy`S>ql_APx_S4Hi$+FcTvrn*;D;p7zGr9boC{2p%8l09nyMBvG&aV0Ltni$-;l&G5>F{c@S~pe85;ccV7591c%B5;>#+'
    'kOdRymG??a;R1|XPr+gZb4EmMKnE}PT})uttwg96!8*EbjD#Xydb=tKfY9J~YrD$`YOqV7jKB_gnx_a6zYkzr<(qs23aZCg%zgDk'
    '!b8?PdDoJs1H`2q1L(_{^8jm|GbTli)HYow`nsZMp<HBjAt?%N8?Pu(z91c4RqyqOmWyb&jj7Af&W8lx>U;GZXa-g`<^zw_3ZHIX'
    '{TER~O|o;~q|=peVxuyarq_6-kg`%(z)RJW8M^L^s^pq#x?umPo<qwOR0lODJQZg{P%{?1X?H39D6Fn~4(P%Qfq~<O*=WUSKi(Hx'
    '@f7oH19&c8hHXSifih5Ob?tLNcd9i}eB1Gxmt{Lp65oK{n0~-ihxbA=)!JB_67$ty!LDtysGH2XS_Ara!{4~w5Uih9DLs!btT#ty'
    'Zu+*&-lI<xEG`CWwbX1V2>{TX4PeGLO3~WQw{D#_b;knb*${Tymc9D@`{HUqDL2fs-Y$l=*nMYt)3lY{S?6G^8WNv+<V{dKZ)=lZ'
    '=`?J_sEXg)R>i4ib-){{lmS9-HK<j^=RfbGX=mlKI%kPwDIc@7RqpTf$vx36EA(oP9L7YkgrI>DXkJ+UWZH|eD&-c#$<ttISu}+^'
    '1a+c)|EH7H8<?lJn6~eleNnJ5)g(tR7qc@7$S^eEwGTcYCPM>_3Es@oPlrdWGHNX_YSZ`SDk-z$GZ0|_R7ZZJn{$cXcOJ}HD%HM('
    '6oefP9!g{L;*I(B?|(UB8?2EnJ;+x#HlbO9bN<L7-kp2MiOMFyKq)(ad{*bb9@O_Q_q%@&txq_{2{_wv(VXI8I_;AdcXwP=5Zh#p'
    '|F6Ak*_Gosvj1||%|^Z-+))(COGqLK(t==b48fy?0Le2Ldw}u6|L!;^A}gZpU*P6;-|p(Htb9bq^BlNbtj^Kc(7~a^6fIW_7nZF^'
    't_6cLA2$q|kj025R!CcbFhk-t-2naaz*H88oV3%o$7o}%!K#762M#8ng28n6vT}G6#B#yMy9`hIYw}}dxPyS3!TEw(VVE{3DoP>0'
    '4~0);vJnx}1B=hrVNHR@r&P{O86z_4iCVIX67SW0o=5Y1dO@~xLgD!^&Pg5!beeHUbg(sVUMJGR%*NieR^lypwlr6ITIwQ4PD?40'
    'W`cf<P4u%CKAxM&5(Vs@;i78O*_y7WX?j#Ji_KPo(h{1+OKud>ZG#0X)Ryf`3K3EC#N-X06C9^x2n60po`JSUbXJ@uiGT^`8(mpx'
    'k$`07!|r0v5U9OP(bDD|SE3$9P>lRloi38He`zyFU5c!na5;cX5v;MRdYSZ|U3rhH5eDgj00<jWjrP+^Jr1C3CluKMQ_6wIFv3Hp'
    'b*WyW^1e6XrQLeFR+ugvEICe?oEtlm=l}|y{)i&|H;OA$p9=sckd2J8;uc~3nMceCVnt24#CHSgJV}IFt+q0LP%LOJqndLp&?5dU'
    '$-`zoICLu17a1!t+S33`Rkw!Eqfoc5xq5tEp23#?xR$?)Z4eDCQ9b|(Y37j9@@(LyDvRvdFWZ(da56C6qXoTt$|*uCYypzqvBDM;'
    '20KyRra{zgBfaNZB|o8OW)BdR$a)!d2jUypVM$CCB>T%Z2dw(t1My8E<W{E1qRQh6bmdm?sF#W~VLdq*k9N?-8#9bqC?A+@fZ_{}'
    'q`Z+e;6+h#Cz2e9Y-oYn2Eqx>KY1hCQ*fKJn3!K6Q*t@84ceLAS~4a?YYQjw;5T!s=3+Tilc{wx82y9kw!>5}yQ3%y@JJwJFy{=M'
    'HekmhMGfluQwiKPqQX>?k`$3r!;778<=FZYnT6mRM}6TUI<CTZ9N#zr^1ap<2+$;%UXeFFE1*orj-1X=Q-Gt8FSJ8@>ZR24&Z>Cw'
    'M^Bcq`M;I>BbyISC$|#`n=mNOLnDt;hEC!vCAhhfoz<E3qbmLxg4u16tzI88#a6w@x)0ZkSO5*4I=q=HBwPb|kX$|EZueBy7a;dR'
    'g9aXE51wpT^|Xu3m@$o?4Bd3ajeAh{3(yD<tAH%sfo=x0C7eHS{_VgvQw!HluI7sUbOvYX^=qpVc1UGpWu%mKi7d-jNoU$xX*zK='
    '90V^=dyJeWI5taQlmU*-;0hnT?7-SyG$|b>UI2DR5iK046?t+cP|5%@QCB)ZV}i3<`Z#tYz$x(c9@4-jN$nNN^V)H1kyAt-Amv~y'
    'z!ISmI;=d<Yz5G^!hPOl5H&J%Caoz5Y`Ti4t9{-L!;Spxfmf!8^d@IZSxF%z49$7a7URM1j%M{=B4a8znrG371OiTd2R;}X7Mn|h'
    'HvV+2=+#V$t+fLkOcBwqYdsGI^?QF!zX-6QBdVL``H~%AF61{k+5o6whk1nKAUnAWHu(`x!&GCF&<i^Z5_+x_D}*pe%0LGErC=Ih'
    'm-Wgodf&4K1{Zn8nib-9DIU22Bu^q$X9KT!do4TzOf)~y5tw2;>Ig0vQhEk`_xY!A1`!B&fpV|huW&AcMmI|~Jf%OVMtNa{4tkj2'
    'CvczFGVkWMGJw|!XG2C&V<@5J<$85@{|!OJrhT3-g3iMg^WY)w%xvu!CS)eQXh-MUJp@-lVm{fiZilErxS~`A5fi|z2Fd}1F7YQM'
    '7as<X$npZBkt<U;E6%LmJC&F!T>m{)IBR~NmCu=@hB2Lbg2>GP+cg&DG-4PutWJ)aJ7+W^h)0AMdW6K?dRKa%8md|-p{EYT<@(To'
    'eZO11n%)fsL%7+ndZ26s9}Q(zIcByl-e=vNVuecyF*eR<5KtJYymX{;fSZT&7=L`r<Nh2|QBY%K2dEHSq@ba_$`#l!xDwfLAX7sB'
    '5lG|DT!CZcZL;}vH-MyK^MmP0p2L`l+1jk><2e2BF%%JyG-MLse5P&0P=|61oJlRvjEvMl^n&<_Vayu*8w1(WtuzD6NCU<nSQt96'
    'ek04s0DfCxJ56t?5g1ri0B#4?R3I4{K%+nj&48BH1IfrJMd?b$+FB0qyiY?#f3gp7_-PAHV16NH{ctpLKCP3Bth=HHj<ZWQi_r%;'
    '97JTAx>2mr0mTiCS}f7ru?1m)#PsprElVWZql)3b(UB>~Q3|xz;BEleK!hFVvDiZ8`dx>wh6(`|@N0#DTGSSSqy@^r=$u-R{{Hl!'
    '65T&(w0E2Mhv@e(QDf|fz#0LG4O=l=o$$$K7))l!%v^VrHPG}B4(8J3I62Z=4wq;{fle0@`f5;YhaExGzKv(?Vk_1WNIpg$XYR#%'
    '>(Q@EGB{8=IGP%s001MV!77Xz>I0i8%9if0m>G0PD^R9}n7T5tJcOg*BCEQVW^Dnz1%-3~iL*$JM(BZ}hy9uJfs7t!?QC7BZLfJL'
    'betaKmUDuXLEeO3(<-rYylrQ7j>_Rck^b`B0{~*n1SSLW)FB$$Muxf`VqS0$0U8905?r8{E{AT4WvF|z_lP|%U96`kw#91G-93S_'
    '>LD^<ccJ9#Zp2U7&`;IhEUR?864#h9Qw?s$^f$n{YXVF+9qaw+U4`<fn<+=y4bb3mm$f!5CEgZRuVBQ(14%C<CMweFjSdA9StG*+'
    'tXW_uw|jTk77PyYqB)70@^CmbI^kMwWL8+sVO!YcU7B?{WDIOXE?E=$fN{zh$TkD5bf?<o5^8CvOj`kU{6(SI2X3@Q_TKNaw1&;l'
    'FiwlZoFN-?VgYo-YuuZ_n~bS4t}Ux&Jgw!K^#Dm0Pd{Bw3nk^)AUPWl(0LfK=uOm*?p}z381Jpj$d7?L$fMh`%sXEJz`%|HGmr>5'
    'cw@TpaYJmY15;L~qL)wlx$#<C55?iY5>y&gtLssQG!zS0klx7|RGYKGLq+RAEX13o{4K`+2e=n>#~2;2%SZ2nm+usE`{cr0z{rd@'
    '&ERzQ)zBKJt>t*+c-9U4F?^#ysK-lz>+urs$R~?u)nI+UPRnao{e2Z_>_z1q#@h6zDb$p|00Wp(r1q8XhVpDrT;OKKYnx>TH^rJv'
    'A%fRDQtZMy<coVMCTN&!cOOSZQzl#hk=B^Rhk(n{e6nTqC)dbNeWo5X+O34ddj{C?4crykY=`#8)<9)F`Or!oJK)+#^rAGNh7VVF'
    '5hu@VVq9k&yFj&yy{IQgB%PkHgx%Pvi?Sk$`i0t8pb3>aT}hS}QdJtom3m9jn*nfL*5$SPj?}rx9fAZyK2mgN6b+UkB2s2_XRJHD'
    '9LbzH$Dgs36HQF~1GcAYW6GISxqIZrkX`J~U71+>aj=}h>o$19Tzp2SC`(0l4cE+U!yOtrg20>6RbN6$TUnZWi7b)j;ZE+L(KdHb'
    'uK@!T#r-%q<P@yIqamZFqc?T*-WxyVs@RE?By|P~4s>z><_q7XbK;*lt#m_8uWaM)+)X#<YN2+VRh&0&i<p4dnWiNRy<hn^Hf>FV'
    'wZuP-4vpOOq)MPWafb#wT4{RXWjd+~<QIKWlxWHF9&7=kYM+uR_Yu%MF{woh^?N^*J?8YtB$tub6hogB#v-VB{}FJ^;P~0l@v(AB'
    'jD4cZa`+5g2`iE~g1L;jo0&?&yv{@%kZq8Pbc!li95@Ay$f2a7$jTedAA&R(S?d$LcNYYwiVO${l^T2cp$I1M4bDoO6?p^*PL3{;'
    'mbp^2fb+fY*=~ClY!}y|h_2RceIC7d(YQjxCLIFMZZC(~!~*WRfW1hZBDzl#(e28MdH#AVemxey9*bX(#jnTW*JJVPv3RI2p1*nE'
    'd#5l&*soQO7Z_i7;G|MM!EyJa4VZ<^_@fn^$mkO9voyA7OpJ%seUU8UE~a+ifJf3l_`PHWr$C}8ZD_AbWKe{iUiD*j(3<F07kxxL'
    'nLIi4h_?HXEFO?l0gw|!Gn2Ppdj*<A!03}Re2$PFS{B7bInc}N@-}C4L7vdYyd3OyNCDcDluYLEw%hDH*r!}Vl@+@_Lc9T?N$leL'
    '3|z#7@<lKp6h<|@(5jP+;!uW?bZ;2Thj?+#=SGLe5Lg%-jS}<KdERC6IF^tdT6EI`7GlSmC&nTXi3kDV#th^3a7<wR$O0#+c&#ng'
    'G?8-$$bmC(69OEb3utCTrb=fJ0U}8xv)~+@yyF2!KZ{LHlN-t7S`_~T206-*KC<TWjUN4iSSBb{!9A0ej2Roe$mXT%C6@=75A6OH'
    '$>=voZP?pDM~_$^@+m};hrlk9S|(GN_u>2Q7B!qjar;)Jq@SHEl)$YG-nwkTcob2t5f}p1@l6r%Kr0WF2+Rsjr${V+Hs8vIvTKXD'
    'A?cNS=7CqJDBU(rr8S)Ek0&9Ye+522B=d4$=eD!ft}cu8f9OmPTz)bDY3rg6k)Ot#hLE;{-mbyx0GVvHG~a&-5)h~sScS*~P`u@L'
    'abE{4eu%CK&`79%sRraIJR`Fg4h^-+JV459(1?nn)k=3HMLEhA?l_05{<sQRniVv@F}j1G;^yVKwvioRv51srBj@T<3zO3315H(M'
    '4>KRZjAqdI5q)k9`&}_u{Je=;k6%YtJ`9s|x^bN3>~Xao{py9IB8tPX`amPC%anSca?36;ttX~{Ws6^K%`bSM*h35%%fFs&c!sZO'
    '2QQ&nbmB`;-7~5rpmU)(Gf?J%5}3?g<sMO)faebvaY0v{AQpq1bvRBjt^yOW!gd;JumPmyu}~L>Sa7L1gmLF0uIUS<6>t$+4a-o('
    'Q~+QC<q&~=y5T09ItPXcD$fgeQ!ZRP^Hb%;ifv820+@!IN|d8)IIDKs!mu)e(Gj<>0GjT`4KW>32}1ztM2^9%C9u{FrYy1?AcURq'
    '_#825E<oqw_^@!4K_6Z$fUqIJ*0enOjKw62o!#kk2KrZ|%O=i$6a6Nj7AnW)h%mpLzNS~tQG&chbqIP>w{ad`q;=3CNm)J5UeQSz'
    'G(A~k6-nzrUR!nDtiX)U5<!ka2kd;LjzcXwW)X&>#R5zaRdBO$hvi*Lz!emM+#$w9pop@#T__{zdgz>4EzY$k(<l@VqVB9my0>gr'
    '@kmxCd&G&zDBV}JxPB|b4mf}nQcu+W7N7zdzN~ymP(=o)z&tzkDp?9^TfOC3OvohVZF-v#c}R_$dErhTu``znmkHj2Bf`J^90BVq'
    'GMxamTr6qVBKPU0u!<_C8<Uv2RS+qzaAC5iyc#vYfsO~6A&OEL0Odvijt4I%5U7#mA3@_0iffB4X|}56y)5g(TU(G&LOr8BlNCxk'
    'Q12-Rr1+u<++v`N!zZI%Zj}kjiwU(}7)hZ4CeSNy9o(Sl-aP#g$Ky$Vvd2f|Hg95Szx(=~i&AE~uM*PB8&l4!eZqnXTw^>-&*#aq'
    'jg+B)ni+!xsk8#1daq0kU%fQT|5mOvYMqr(3=Wx|6+Hs91D+EjCIkqyF$b}9eIeIR2e%&w2ro}4)V-C%kam#DJClepwA)yhtymI|'
    'TS`a=N0B8`@q0)GN3DX+Og(wp2KWtXd1SgSrdG8OG?uH6Z^4!kNHp2wbsk$WyE(fv5@}+ouiIXh%W&>BwTKZW+1)C2b(55>y+fT&'
    '3eegu>F}e-ETCs&VRl=XJ>X*_X@{4mgJ)s0IEgcvK_Grs@|lHdEtyGalu2CpGm<3{8&x~mF@(lW@yE{(*w<qHcPwO4i(63FG}kxe'
    'xPi?Z3ADpEJTa|-t9sDNvq8E-g(ZOO+3C0v);zD3MVPO>j&DT!=NCn@TrMCHC>925c>#%ZGO4bBUoH0i)K18AI`$?XEfZ@vr}Y3-'
    '4T?;By;2Dr#+4KdC=#1mS9?K%{3jqt0v<>T(9?TQc46dTdQxG3RHMMK-4!>{Y>;YhO`M;*2O+@Z4&;aAwp<c$B<D4f(FV;_L+*}X'
    'shHk)RKsPqEDb6|gUn{Z0M&Y3zNTP-a&NxydL6X~7j?w8&ZpXhwu~3ogfQN{Ep{VhqTo3|&c)}LBV-nXh7}6ffuFw?)l|gf0DysQ'
    'y$K+ig0Z<&kJ>4S6Wkn${9FyKPEROCi+L(O`vdkx=4jT4+Pn$mo)xnF*qY=4l$ZU8Z$Z91CZeQ>&}<K}J(m|<%Wz$wYqdw1<Ei^-'
    'AKp(}Y+X!A!WpkzobZiXAy?Z9xwx3&ZAJZOTgweXxNr0hy&Uu8%uTL3HDv@|*j5>?Y0)U;z~!w2)9OLkfBqE8s8LQZZY0E>A&nL;'
    'LlevD`^$9VWiNxKD*!*#>@BT90(<NYf8E73Dw9EQ<yhS-F8Dt_wW>?Yo&}zzzFe~9bmu}S;*Z3P@*+vEl$NiWHaol!x&dt#{n8q)'
    'x8@sdxK(SeMx-PP>qDzo5-odm(Vn97S#P6U2dDR$Jx$5WCV)c<tpDyVE$n8FrfqV&48Bm+hV3B-$9GIOy$0%&JDx;YwcnYSb<+>G'
    'y}$ur!t%+LrKUdbT6taV&Qu0xF1^s6%?4n8Dpa*O<rql+>WzP!TSj2gg$v!w3<f}jIMzM8VuAAzkC&Fm9T1eG&(=c@s1crADm!qV'
    ';9r*d<JD~*PIcC)p4X;#o$jqYG-?rU4`{G>3W&YY4zB367Yc8)s+Z^@_ZO6`MCkD%T=}0LmMhC@Sd|)H>vp7Z?k#pcc}}nV6gbH8'
    'GLt7mA#q&O$33MN&26^24vW?6kDsDvoAlzhN(-+A*ET%$fgS@yM_`ixK5+9&tO^j+TQ!6!qsAD6r(T;#&vHM{n4t9q=8TecnnP!7'
    '=BJD3HoPnI=d?wrk4iS%7Tnv31`2q;Ebe1A0~n|e8jE_hz;*>+=uDbeZBQ!EnL)mbjCiXlp*aYGKKLh_Cl^jD<V58QnO%Ilw{rIX'
    'yXzRandDH{Ohwt_4S$P6fl=4)Munw2x^^bbwv%(o%b-^Fc~0_l;tl!kF-_6$$oN|U^Wlm>Vbt&jo6TzVhZZIyw{Av~h(#8SbwU`*'
    '#s<^9+~)P^G4&vvp)>?0;JON#nLy-lt{%qswG?_KQa-fAK&yy=3oAg~lrSjZ8|@ddh2E3wl6mDw-dkaHz+DI?<l(%%lh+OnWt#cX'
    '84@TxA7fd)sDso+Na(!J9o*AP7b0k;Oa$w&v+QZi;GWKj3#A_4gqa^oERsQo<^E}gm9e}}5nxv)b20qqH0PR*JWLT8dFW}uMN(Oq'
    'i@=yJRwH^C(tDio1BAd{x|*wzZP#f4HR@4;@&&3uz(kbb-3Hzjh8~sa?n(iw#01~hIPC0a1{Y29*3Fmoe0C1&Q~>K_dC+=OaH@b!'
    '+jgqaDu9&}iY~Y$Y7g#hMf_<OxU{WuV0B}a=MpWZ3M>aKJ_7P=H<$)_(}pr29|yNA)Q(WuWYL63!sq|~?brYM`5%81S@=)+_xE4_'
    '`scs?-|xTt@2`J>bo}Af`wveq)9vZix1V3Udh>?w-~8e2(@XHTJ$>`>!$<iH6^T!;e~@=sp8n~}Pp>|Gd3q_k^y&5YZ(n~ea#qe!'
    'y#N2#uYS-be*FCW-M3FKWpL&5Pw{gR%}+1ooO^op<G0uQci()v-hcR$e-1g0_J_WI_4x~WqZ-;9>3L6l|K{DN>+?Uo68lh&YyS3?'
    'xATXe<bBuAzr6YG`uzFFx7YRT^SS=>AFubM=bg#f5%FAyhsoZPgu<=|9JakB#HPI`Y*!nO(_`Hs_=7zni-_%3Eas}q%Tvx{eef<b'
    'ugTssnA?|!HLOkg$Lzi|pKaKv<hJI=C4TD*CdqpF`!B!$&u>5f{g*%gO_t2hzoD7h+CTpVrd4GV'
)
BASE = json.loads(zlib.decompress(base64.b85decode(BASELINE_B85)).decode())

def parse_date(s):
    for fmt in ('%Y%m%d','%Y-%m-%d','%d-%m-%Y','%d/%m/%Y'):
        try: return dt.datetime.strptime(s.strip(),fmt).date()
        except ValueError: pass
    raise ValueError('Fecha inválida: '+s+'; use AAAAMMDD.')

def serial(d): return (d-EPOCH).days

def sha(path):
    with open(path,'rb') as f: return hashlib.sha256(f.read()).hexdigest()

def colname(n):
    s=''
    while n: n,r=divmod(n-1,26); s=chr(65+r)+s
    return s

def number(v): return isinstance(v,(float,int)) and not isinstance(v,bool) and math.isfinite(v)

def same(a,b):
    if a in ('',None) and b in ('',None): return True
    if number(a) and number(b): return math.isclose(a,b,rel_tol=1e-11,abs_tol=1e-10)
    return a==b

def normalize(v):
    if isinstance(v,int) and v in COM_ERRORS: return COM_ERRORS[v]
    if v=='': return None
    if isinstance(v,dt.datetime): return serial(v.date()) + (v.hour*3600+v.minute*60+v.second)/86400
    return v

def read_xlsx(path):
    """Lectura de OOXML sin abrir ni alterar el archivo en Excel."""
    with zipfile.ZipFile(path) as z:
        w=ET.fromstring(z.read('xl/workbook.xml'))
        rels={x.get('Id'):x.get('Target') for x in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        ss=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ss=[''.join(t.text or '' for t in si.iterfind('.//m:t',N)) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        data={}; formulas=0
        for sh in w.find('m:sheets',N):
            target=rels[sh.get(REL)]
            target=target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target))
            if not target.startswith('xl/worksheets/'): raise ValueError('Tipo de hoja no esperado: '+target)
            root=ET.fromstring(z.read(target)); cells={}
            for c in root.findall('.//m:c',N):
                f=c.find('m:f',N); vv=c.find('m:v',N); t=c.get('t','n')
                if f is not None: formulas+=1
                v=vv.text if vv is not None else None
                if t=='inlineStr': v=''.join(x.text or '' for x in c.findall('.//m:t',N))
                elif v is not None:
                    if t=='s': v=ss[int(v)]
                    elif t=='n': v=float(v)
                    elif t=='b': v=bool(int(v))
                if v not in (None,''): cells[c.get('r')]=v
            data[sh.get('name')]=cells
        return data,formulas,any(n.endswith('vbaProject.bin') for n in z.namelist())

def find_main(folder):
    def key(s):return re.sub(r'[^a-z0-9]','',s.lower())
    candidates=[p for p in folder.glob('*.xlsm') if not p.name.startswith('~$') and key(p.stem)=='ptsfwdandetfv621']
    if len(candidates)!=1: raise RuntimeError(f'Se esperaba exactamente un MAIN en {folder}. Encontrados: {candidates}. Use --main con la ruta completa.')
    return candidates[0]

def business_days(start,end,source):
    holidays=source['Holidays']; available=set(); excluded=set()
    for k,v in holidays.items():
        if re.fullmatch(r'A\d+',k) and number(v) and str(holidays.get('B'+k[1:],'')).strip().upper()=='CHILE':
            d=EPOCH+dt.timedelta(days=int(v)); available.add(d.year); excluded.add(d)
    if any(y not in available for y in range(start.year,end.year+1)): raise ValueError('Holidays no contiene todos los años solicitados.')
    days=[]; skipped=[]
    for i in range((end-start).days+1):
        d=start+dt.timedelta(days=i)
        if d.weekday()>=5: continue
        if d in excluded: skipped.append(d)
        else: days.append(d)
    return days,skipped

def is_error(v): return isinstance(v,str) and v.strip().startswith('#')

class BloombergUnavailable(RuntimeError):
    """Detener el lote: el complemento no responde o Excel no reconoce las funciones."""

def is_name_error(v):
    return isinstance(v,str) and v.strip().upper() in ('#NAME?', '#¿NOMBRE?', '#NOMBRE?', '¿NOMBRE?')

def require_no_name_errors(data):
    bad=[f'{n}!{k}' for n,cells in data.items() for k,v in cells.items() if is_name_error(v)]
    if bad:
        raise BloombergUnavailable(f'Excel no reconoce fórmulas (#NAME? / #¿NOMBRE?): {len(bad)} celdas. Ejemplos: '+', '.join(bad[:12]))

def log_addins(app):
    # Diagnóstico de solo lectura: no instalar, registrar ni desconectar complementos.
    for collection in ('AddIns','COMAddIns'):
        try:
            for item in getattr(app,collection):
                name=str(getattr(item,'Name','') or getattr(item,'Description','') or getattr(item,'ProgId',''))
                if any(x in name.lower() for x in ('bloomberg','blp','bof')):
                    state=getattr(item,'Installed',None) if collection=='AddIns' else getattr(item,'Connect',None)
                    LOG.info('%s: %s | activo=%s',collection,name,state)
        except Exception as ex:LOG.info('Diagnóstico %s no disponible: %s',collection,ex)

def check_bloomberg(app,probe,pythoncom,timeout=90):
    # Una petición independiente, en un libro descartable, comprueba que BDP está registrado.
    # No se toca el MAIN, ni se alteran las opciones del complemento.
    cell=probe.Worksheets(1).Range('A1')
    cell.ClearContents()
    cell.Formula='=BDP("EURUSD Curncy","PX_LAST")'
    start=time.monotonic(); last=None
    while time.monotonic()-start<timeout:
        pythoncom.PumpWaitingMessages()
        try:
            last=normalize(cell.Value2)
            if number(last) and last>0:
                LOG.info('Prueba Bloomberg BDP: OK (precio numérico).')
                return
            if is_name_error(last) and time.monotonic()-start>=5:
                raise BloombergUnavailable('La sesión Excel devuelve #¿NOMBRE? en BDP de prueba. No se abrirá otro MAIN.')
        except BloombergUnavailable:raise
        except Exception as ex:LOG.debug('Excel ocupado en prueba BDP: %s',ex)
        time.sleep(2)
    raise BloombergUnavailable(f'BDP no devolvió precio numérico en {timeout}s; último valor: {last!r}. Revisar Bloomberg y permisos.')

def is_pending(v):
    s=str(v).lower()
    return any(x in s for x in ('requesting data','getting_data','getting data','#busy!','retrieving data','loading data'))

def historical_values(data,n): return {k:data.get(n,{}).get(k) for k in BASE['historical'][n]}

def validate(data,day,source,previous=None,strict=False):
    """Devuelve bloqueos, advertencias y detalle de cada incidencia."""
    blocks=[]; warnings=[]; incidents=[]
    if list(data)!=BASE['names']: return ['Hojas u orden distintos de la plantilla.'],[],[]
    s=serial(day)
    for n,cells in data.items():
        for k in BASE['dates'][n]:
            if not same(cells.get(k),s): blocks.append(f'{n}!{k}: fecha {cells.get(k)!r}; esperada {s}.')
        for k,expected in BASE['static'].get(n,{}).items():
            if not same(cells.get(k),expected): blocks.append(f'{n}!{k}: identificador distinto al de referencia.')
        for k,v in cells.items():
            if not (is_error(v) or is_pending(v)):continue
            known=v in BASE['known'][n].get(k,[])
            pending=is_pending(v)
            status='CONOCIDA' if known else 'NUEVA'
            incidents.append({'hoja':n,'celda':k,'valor':v,'estado':status,'pendiente':pending})
            # Un N/A de mercado nuevo puede ser legítimo; exigir cobertura numérica.
            allowed_market_na=str(v).strip().upper()=='#N/A N/A'
            if (pending and (strict or not known)) or (not known and not allowed_market_na):
                blocks.append(f'{n}!{k}: {v} ({status})')
        h=historical_values(data,n); min_count=BASE['min_numeric'][n]
        count=sum(number(v) for v in h.values())
        # 90% de la menor cobertura de los dos ejemplos. No se imputa ningún precio.
        if min_count and count<math.ceil(min_count*.90):
            blocks.append(f'{n}: cobertura numérica insuficiente ({count}/{min_count} de referencia).')
        # Verificar que precios históricos hayan cambiado, no solo la fecha de cabecera.
        for tag,old in [('caché MAIN',source),('reporte anterior',previous)]:
            if not old or same(old.get('MAIN',{}).get('D4'),s):continue
            old_h=historical_values(old,n)
            variable=[k for k in h if all(number(v) for v in BASE['refs'][n][k]) and not same(*BASE['refs'][n][k]) and number(old_h.get(k)) and number(h.get(k))]
            if len(variable)>=2 and all(same(h[k],old_h[k]) for k in variable):
                blocks.append(f'{n}: {len(variable)} precios variables idénticos al {tag}; posible dato sin refrescar.')
    if incidents:
        warnings.append(f'{len(incidents)} incidencias Excel/Bloomberg; {sum(i["pendiente"] for i in incidents)} pendientes, {sum(i["estado"]=="NUEVA" for i in incidents)} nuevas. Ver JSON.')
    warnings.append('La plantilla incluye BDP sin fecha, TODAY(), BDH relativo y USDCNF vinculado a USDCHF; se conserva la metodología original.')
    return blocks,warnings,incidents

def compare_reference(data,day):
    idx={dt.date(2026,11,17):0,dt.date(2026,11,18):1}.get(day)
    if idx is None:return None
    diff=[]; count=0
    for n,items in BASE['refs'].items():
        for k,vals in items.items():
            expected=vals[idx]; actual=data[n].get(k);count+=1
            if not same(expected,actual):diff.append({'hoja':n,'celda':k,'referencia':expected,'resultado':actual})
    return {'celdas_historicas_comparadas':count,'diferencias':diff}

def compare_saved(snapshot,saved):
    if list(snapshot)!=list(saved):return ['Hojas distintas después de guardar.']
    diff=[]
    for n in snapshot:
        for k in snapshot[n].keys()|saved[n].keys():
            if not same(snapshot[n].get(k),saved[n].get(k)):
                diff.append(f'{n}!{k}: {snapshot[n].get(k)!r} -> {saved[n].get(k)!r}')
    return diff

def snapshot_excel(wb):
    result={}
    for ws in wb.Worksheets:
        name=str(ws.Name); ur=ws.UsedRange
        rows=int(ur.Row+ur.Rows.Count-1); cols=int(ur.Column+ur.Columns.Count-1)
        if rows*cols>500000:raise RuntimeError(f'{name}: rango usado anormal ({rows}x{cols}).')
        vals=ws.Range(ws.Cells(1,1),ws.Cells(rows,cols)).Value2
        if not isinstance(vals,tuple):vals=((vals,),)
        cells={}; headers=[colname(i) for i in range(1,cols+1)]
        for i,row in enumerate(vals,1):
            if not isinstance(row,tuple):row=(row,)
            for j,v in enumerate(row):
                v=normalize(v)
                if v not in (None,''):cells[headers[j]+str(i)]=v
        result[name]=cells
    return result

def wait_for_bloomberg(app,wb,day,source,previous,args,pythoncom):
    start=time.monotonic(); last=None; stable_since=start; last_log=0; last_blocks=[]
    while time.monotonic()-start<args.timeout:
        pythoncom.PumpWaitingMessages()
        try:
            data=snapshot_excel(wb)
            require_no_name_errors(data)
            blocks,warnings,incidents=validate(data,day,source,previous,args.strict)
            last_blocks=blocks
            # Comparar solo áreas históricas y errores: BDP en tiempo real puede moverse.
            sig=tuple((n,k,data[n].get(k)) for n in data for k in BASE['historical'][n]+BASE['dates'][n])
            sig+=tuple((i['hoja'],i['celda'],str(i['valor'])) for i in incidents)
            now=time.monotonic(); pending=sum(i['pendiente'] for i in incidents)
            if sig!=last: last=sig; stable_since=now
            elapsed=now-start
            if now-last_log>=20:
                LOG.info('%s | %.0fs | pendientes=%s | bloqueos=%s | estable %.0fs',day,elapsed,pending,len(blocks),now-stable_since);last_log=now
            min_wait=args.min_wait if not pending else max(args.min_wait,min(90,args.timeout-20))
            if not blocks and elapsed>=min_wait and now-stable_since>=20 and int(app.CalculationState)==0:
                return data,warnings,incidents
        except BloombergUnavailable:raise
        except Exception as e:
            last_blocks=[str(e)];last=None;stable_since=time.monotonic()
            LOG.warning('Excel temporalmente ocupado: %s',e)
        time.sleep(5)
    raise RuntimeError('Bloomberg no pasó el control antes del límite. '+ ' | '.join(last_blocks[:15]) + (f' (+{len(last_blocks)-15})' if len(last_blocks)>15 else ''))

def freeze_and_save(app,wb,temp_path):
    # Equivalente a Cells.Copy / PasteSpecial xlPasteValues de la macro; sin MsgBox.
    # No guardar el libro .xlsm. Se trabaja sobre su apertura de solo lectura.
    for ws in wb.Worksheets:
        rng=ws.UsedRange
        if rng.HasFormula is False:continue
        rng.Copy();rng.PasteSpecial(Paste=-4163);app.CutCopyMode=False
    old_alerts=app.DisplayAlerts
    try:
        app.DisplayAlerts=False # Solo durante SaveAs a XLSX (elimina VBA de la copia).
        wb.SaveAs(Filename=str(temp_path),FileFormat=51,AddToMru=False)
    finally:
        app.DisplayAlerts=old_alerts

def publish(temp,target):
    # Copia temporal al mismo directorio de destino; en Windows rename no sobrescribe.
    target.parent.mkdir(parents=True,exist_ok=True)
    scratch=target.parent/('.'+target.stem+'.'+uuid.uuid4().hex+'.tmp')
    try:
        shutil.copyfile(temp,scratch)
        if sha(temp)!=sha(scratch):raise IOError('La copia no coincide con el XLSX validado.')
        if target.exists():raise FileExistsError(str(target))
        os.rename(scratch,target)
    finally:
        if scratch.exists():scratch.unlink()

def setup_logger(folder):
    folder.mkdir(parents=True,exist_ok=True)
    LOG.setLevel(logging.INFO); LOG.handlers.clear()
    fmt=logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
    for h in (logging.StreamHandler(sys.stdout),logging.FileHandler(folder/'ejecucion.log',encoding='utf-8')):
        h.setFormatter(fmt); LOG.addHandler(h)

def run_day(app,pythoncom,main,day,source,previous,target,folder,args):
    original_hash=sha(main); wb=None; saved_data=None; audit={'fecha':str(day),'destino':str(target),'intentos':[]}
    for attempt in range(1,args.retries+2):
        record={'intento':attempt};audit['intentos'].append(record)
        try:
            with tempfile.TemporaryDirectory(prefix='PTS_') as stage:
                tmp=Path(stage)/(day.strftime('%Y%m%d')+'.xlsx')
                # Mantener la política corporativa de macros y la sesión Excel existentes.
                # UpdateLinks=0 solo evita los cuatro vínculos heredados a libros LIMITES/EXCESOS.
                wb=app.Workbooks.Open(Filename=str(main),UpdateLinks=0,ReadOnly=True,AddToMru=False,IgnoreReadOnlyRecommended=True)
                if not wb.ReadOnly:raise RuntimeError('El MAIN no se abrió como solo lectura.')
                if [str(ws.Name) for ws in wb.Worksheets]!=BASE['names']:raise RuntimeError('La plantilla no tiene las 46 hojas de referencia.')
                wb.Activate()
                wb.Worksheets('MAIN').Activate()
                wb.Worksheets('MAIN').Range('D4').Value2=serial(day)
                # La edición de D4 activa el recálculo normal de Excel. No forzar
                # CalculateFullRebuild ni llamar macros internas de Bloomberg.
                if int(app.Calculation)!=-4105:raise BloombergUnavailable('Excel no está en cálculo automático. Active Fórmulas > Opciones de cálculo > Automático.')
                snap,warnings,incidents=wait_for_bloomberg(app,wb,day,source,previous,args,pythoncom)
                # Capturar los últimos valores sin cambiar el modo de cálculo.
                snap=snapshot_excel(wb)
                require_no_name_errors(snap)
                blocks,warnings,incidents=validate(snap,day,source,previous,args.strict)
                record.update(advertencias=warnings,incidencias=incidents,bloqueos=blocks)
                if blocks:raise RuntimeError('Control final previo al guardado: '+' | '.join(blocks[:12]))
                freeze_and_save(app,wb,tmp)
                wb.Close(SaveChanges=False);wb=None
                saved_data,nf,has_vba=read_xlsx(tmp)
                require_no_name_errors(saved_data)
                if nf or has_vba:raise RuntimeError(f'XLSX final contiene {nf} fórmulas o VBA={has_vba}.')
                mismatches=compare_saved(snap,saved_data)
                record['diferencias_congelado']=mismatches
                if mismatches:raise RuntimeError(f'{len(mismatches)} celdas cambiaron durante el pegado/guardado; no publicar. '+str(mismatches[:5]))
                blocks,warnings,incidents=validate(saved_data,day,source,previous,args.strict)
                if blocks:raise RuntimeError('Control del XLSX final: '+' | '.join(blocks[:12]))
                ref=compare_reference(saved_data,day);record['comparacion_referencia']=ref
                if ref is not None:
                    LOG.info('Comparación histórica %s: %d celdas; %d diferencias.',day,ref['celdas_historicas_comparadas'],len(ref['diferencias']))
                    if ref['diferencias']:
                        warnings.append('Diferencias respecto del reporte de referencia; consultar JSON.')
                        if not args.test:raise RuntimeError('Los precios históricos difieren del reporte de referencia de esa fecha. No se publica; revise el JSON o ejecute --test.')
                if sha(main)!=original_hash:raise RuntimeError('El MAIN cambió en disco durante el proceso; se detiene la publicación.')
                publish(tmp,target)
                record['resultado']='GENERADO_CON_ADVERTENCIAS' if warnings else 'GENERADO'
                record['advertencias']=warnings
                audit['resultado']=record['resultado'];audit['sha256_xlsx']=sha(target)
                LOG.info('%s | %s | %s',day,audit['resultado'],target)
                return audit,saved_data
        except Exception as e:
            record['error']=str(e);LOG.error('%s | intento %s: %s',day,attempt,e)
            if wb is not None:
                try:wb.Close(SaveChanges=False);wb=None
                except Exception as close_error:raise RuntimeError('No fue posible cerrar el libro de trabajo; se detiene para proteger Excel.') from close_error
            if sha(main)!=original_hash:raise RuntimeError('El MAIN cambió en disco. Se detiene todo el proceso.') from e
            if isinstance(e,BloombergUnavailable):
                audit['resultado']='BLOOMBERG_NO_DISPONIBLE'
                return audit,None
            if target.exists():break
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False);wb=None
            app.CutCopyMode=False
            # La sesión Excel sigue abierta y con sus opciones originales.
        if attempt<=args.retries:time.sleep(8)
    audit['resultado']='OMITIDO_EXISTENTE' if target.exists() else 'ERROR_NO_PUBLICADO'
    return audit,None

def args_cli():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--main',type=Path,help='Ruta completa del XLSM si no está en Maria.')
    p.add_argument('--output',type=Path,default=OUTPUT_ROOT,help='Raíz Pts FWD & ETF - Bloomberg (sin año).')
    p.add_argument('--start',help='AAAAMMDD; si no se indica, se pregunta.')
    p.add_argument('--end',help='AAAAMMDD; si no se indica, se pregunta.')
    p.add_argument('--test',action='store_true',help='Genera en carpeta local de pruebas, no en SharePoint.')
    p.add_argument('--strict',action='store_true',help='Bloquea incluso Requesting Data conocido en los ejemplos.')
    p.add_argument('--timeout',type=int,default=300,help='Espera máxima por intento (segundos).')
    p.add_argument('--min-wait',type=int,default=45,help='Espera mínima antes de congelar (segundos).')
    p.add_argument('--retries',type=int,default=1,help='Reintentos adicionales por fecha.')
    p.add_argument('--diagnostico',action='store_true',help='Solo prueba BDP en la sesión Excel existente; no abre el MAIN ni crea reportes.')
    p.add_argument('--plan',action='store_true',help='Solo muestra las fechas; no abre Excel ni crea reportes.')
    return p.parse_args()

def main():
    args=args_cli()
    main_path=(args.main or find_main(MAIN_DIR)).resolve(strict=True)
    if main_path.suffix.lower()!='.xlsm':raise ValueError('Se requiere el MAIN original .xlsm.')
    source,nf,vba=read_xlsx(main_path)
    if list(source)!=BASE['names'] or not nf or not vba:raise RuntimeError('El MAIN no coincide con la estructura original o perdió las fórmulas/VBA.')
    if args.diagnostico:
        a=b=DEFAULT_START
    elif args.test:
        a=args.start or '20261117';b=args.end or '20261118'
    else:
        a=args.start or input(f'Fecha inicial [{DEFAULT_START}]: ').strip() or DEFAULT_START
        b=args.end or input(f'Fecha final   [{DEFAULT_END}]: ').strip() or DEFAULT_END
    start,end=parse_date(a),parse_date(b)
    if start>end:raise ValueError('La fecha inicial debe ser anterior o igual a la final.')
    if end>=dt.date.today():raise ValueError('El rango debe terminar antes del día actual para evitar datos intradía.')
    if args.min_wait<20 or args.timeout<max(110,args.min_wait+25) or args.retries<0:raise ValueError('Parámetros de espera inválidos.')
    days,holidays=business_days(start,end,source)
    run_id=dt.datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:6]
    folder=main_path.parent/'PTS_FWD_Auditoria'/run_id
    if args.plan:
        print('MAIN:',main_path,'\nDías hábiles:',len(days),'\nFeriados excluidos:',holidays)
        for d in days: print(d.strftime('%Y%m%d'))
        return
    setup_logger(folder)
    output=folder/'PRUEBA' if args.test else args.output
    LOG.info('MAIN=%s | destino=%s | modo=%s | días=%s | feriados=%s',main_path,output,'TEST' if args.test else 'PRODUCCIÓN',len(days),holidays)
    if not args.test and not args.diagnostico and not output.is_dir():raise RuntimeError('No existe la raíz local sincronizada de SharePoint. Revise la ruta; no se creará una ruta alternativa.')
    if not days:return
    todo=[d for d in days if not (output/str(d.year)/(d.strftime('%Y%m%d')+'.xlsx')).exists()]
    if not todo and not args.diagnostico:
        LOG.info('Todos los archivos ya existen; no se sobrescribe ninguno.')
        with open(folder/'resumen.csv','w',newline='',encoding='utf-8-sig') as f:
            writer=csv.DictWriter(f,fieldnames=['fecha','estado','destino'],delimiter=';');writer.writeheader()
            for d in days:writer.writerow({'fecha':d.strftime('%Y%m%d'),'estado':'OMITIDO_EXISTENTE','destino':str(output/str(d.year)/(d.strftime('%Y%m%d')+'.xlsx'))})
        return
    try:import pythoncom; import win32com.client
    except ImportError:raise RuntimeError('Falta pywin32. Ejecute: python -m pip install pywin32')
    pythoncom.CoInitialize()
    app=None;probe=None;rows=[];previous=None;main_hash=sha(main_path)
    try:
        try:app=win32com.client.GetActiveObject('Excel.Application')
        except Exception as e:raise RuntimeError('Abra Excel de escritorio normalmente, con la pestaña Bloomberg habilitada y sesión iniciada. No abra el MAIN.') from e
        # Usar solamente la sesión que el usuario abrió con Bloomberg habilitado.
        if not app.Visible:raise RuntimeError('La instancia encontrada no es visible. Abra Excel de escritorio normalmente.')
        if not args.diagnostico:
            for wb in app.Workbooks:
                if not wb.IsAddin and str(wb.Path):raise RuntimeError('Cierre los libros de trabajo guardados en Excel (incluido MAIN), conservando Excel abierto con Bloomberg.')
                if not wb.IsAddin and not wb.Saved:raise RuntimeError('Hay un libro sin guardar. Guárdelo y ciérrelo antes de ejecutar.')
        if int(app.Calculation)!=-4105:raise RuntimeError('Active Fórmulas > Opciones de cálculo > Automático antes de ejecutar.')
        if not app.EnableEvents:raise RuntimeError('Excel tiene eventos deshabilitados. Reinicie Excel normalmente antes de ejecutar.')
        log_addins(app)
        # El libro de prueba permanece abierto todo el lote; no se cierra Excel.
        probe=app.Workbooks.Add()
        check_bloomberg(app,probe,pythoncom)
        if args.diagnostico:
            LOG.info('Diagnóstico OK. No se ha abierto ni modificado el MAIN.')
            return 0
        for i,d in enumerate(days,1):
            target=output/str(d.year)/(d.strftime('%Y%m%d')+'.xlsx')
            LOG.info('[%d/%d] %s',i,len(days),d)
            if sha(main_path)!=main_hash:raise RuntimeError('El MAIN cambió durante el lote; ejecución detenida.')
            if target.exists():
                audit={'fecha':str(d),'resultado':'OMITIDO_EXISTENTE','destino':str(target),'intentos':[]}
                LOG.info('Ya existe; no se modifica: %s',target)
            else:
                audit,generated=run_day(app,pythoncom,main_path,d,source,previous,target,folder,args)
                if generated is not None:previous=generated
            (folder/(d.strftime('%Y%m%d')+'.json')).write_text(json.dumps(audit,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
            rows.append({'fecha':d.strftime('%Y%m%d'),'estado':audit['resultado'],'destino':str(target),'detalle':str(folder/(d.strftime('%Y%m%d')+'.json'))})
            with open(folder/'resumen.csv','w',newline='',encoding='utf-8-sig') as f:
                writer=csv.DictWriter(f,fieldnames=['fecha','estado','destino','detalle'],delimiter=';');writer.writeheader();writer.writerows(rows)
            if audit['resultado']=='BLOOMBERG_NO_DISPONIBLE':
                LOG.error('Lote detenido para proteger los reportes. Revise el complemento; no se procesarán las fechas restantes.')
                return 3
            if audit['resultado']=='ERROR_NO_PUBLICADO':
                check_bloomberg(app,probe,pythoncom)
    finally:
        if probe is not None:
            try:probe.Close(SaveChanges=False)
            except Exception as e:LOG.warning('No se pudo cerrar el libro de diagnóstico: %s',e)
        pythoncom.CoUninitialize()
    LOG.info('MAIN intacto (SHA256): %s',sha(main_path)==main_hash)
    LOG.info('Resumen: %s',folder/'resumen.csv')
    failures=[r for r in rows if r['estado']=='ERROR_NO_PUBLICADO']
    if failures:LOG.error('%d días no publicados. Reejecute el mismo rango para intentar los faltantes.',len(failures));return 2
    return 0

if __name__=='__main__':
    try:sys.exit(main() or 0)
    except KeyboardInterrupt:print('\nProceso interrumpido. Los reportes publicados no se modifican.');sys.exit(130)
    except Exception as exc:LOG.exception('Proceso detenido: %s',exc);sys.exit(1)
