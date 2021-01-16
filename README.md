# Excel_Kod_Uret
Excelde Kendinize Özel Değer Üretmek İçin Kullanabilirsiniz.
Excel VBA ile hazırladım.
Ben Maillerde Konu Takibi yapmak için kullanıyorum.

Excel Sayfamızı aşağıdaki gibi Hazırladıktan Sonra Geliştirici Sekmesinden veya Buton üzerin çift tıklayarak kodları ekleyip kullanamabilirsiniz.
Dosyanın çalışması için Excel Dosyanızı .xlsm uzantılı olarak kaydetmeyi unutmayın

a1 hücresine kendinizce anlamlı bir ifade (kurum adı,proje adı vs.) yazabilirsiniz.
b1 hücresinde =RASTGELEARADA(1000;9999) formülü yer alıyor
c1 hücresinde =UNICODEKARAKTERİ(D1)&UNICODEKARAKTERİ(E1)&UNICODEKARAKTERİ(F1)&UNICODEKARAKTERİ(G1) formülü yeralmakta

D1 hücresinde =RASTGELEARADA(65;86)
E1 hücresinde =RASTGELEARADA(65;86)
F1 hücresinde =RASTGELEARADA(65;86)
G1 hücresinde =RASTGELEARADA(65;86)

=rastgelearada(65;86) formülü ile amaçlanan A-Z arasında bir karakter oluşturmak amacı ile unicodekarakteri() formuülüne bir referans oluşturmak için kullandım.


