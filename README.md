# <img src="https://img.shields.io/badge/c%23%20-%23239120.svg?&style=for-the-badge&logo=c-sharp&logoColor=white"/> **CSICorp INTELLIGENT SYSTEMS**

<table  border = '0' style="width: 100%;">
  <tr>
    <td>
        <img align="left" width="116" height="116" src="./favicon.ico" >
    </td>
    <td> 

## CSICorp INTELLIGENT SYSTEMS
### Костадин Караиванов

</tr>
</table>
<br />

# Демо на система и функционалност относно обработка и представянето на структурирани данни  от системата Trimble 4D
<br />

## **Общ преглед**

> ***Демо версия на приложение за генериране на отчет във формат xlsx, необходим за анализ на данните и изготвяне на седмична справка. Приложението обработва данни от системата Trimble 4D, експортнати във формат csv. Експортите от всеки ден са поставени в архив, който се подава на приложението за изчисляване.***
> <br />
> 
> ***Предстои добавяне на функционалност за обработване и сравнение на динните от предишен период, необходими при анализа и формиране на седмичната справка.***
---

## <br />**Спецификация**

> #### Приложението е базирано на [.NET Core](https://en.wikipedia.org/wiki/.NET_Core) фреймуорк и [WASM](https://en.wikipedia.org/wiki/WebAssembly) Blazor технологията. <br /> 
> * #### За frondEnd, като програмен език са използвани HTML, CSS и JS, но на този етап от проекта не се използват Cookies, Local и Session storage.
> * #### За backhend, като програмен език е използван C#, който реализира функционалността по изчисление на входните данни и генериране на отчет във файла на [Microsoft Office](https://www.microsoft.com).
---

## <br />**Функционалност**
> 
#### __*От началната страница кликнете на бутона "Reports"*__
> <img src="./images/firstPage.JPG">
> <br />
> 
#### <br />__*На страница "Reports" кликнете на бутона "Избери файл". За демото може да свалите файл от посочените линкове по-долу.*__
> <img src="./images/seccondPage.JPG">
> <br />
> 
#### <br />__*След избор на файл /zip архив/ ще бъдете препратени към изглед с информация за всеки файл намиращ се в архива, както и данните от всеки файл, които предстои да бъдат изчислени.*__
> <img src="./images/seccondPage1.JPG">
> <br />
> 
### <br />__*Най-отдолу във формата кликнете на бутона "Преглид в таблица". Бутонът "Изтрии всичко" ще Ви върне на предишна стъпка за избор на нов файл с данни.*__
> <img src="./images/seccondPage2.JPG">
> <br />
> 
### <br />__*На генерирания предварителен отчет, най-отдолу се намират три бутона.*__
> <img src="./images/seccondPage3.JPG">
> <br />
> 
### <br />__*Бутонът "Избери предходен период" служи за добавяне на zip архив с файлове от предходен период. Бутонът "Изтрии всичко" ще Ви върне на предишна стъпка за избор на нов файл с данни. Бутонът "Създай отчет" ще ви препрати към съгласие за създаване на отчет. Файл ще се генерира дори без добавяне на zip архив за предходен период.*__
> <img src="./images/seccondPage4.JPG">
> <br />

### <br />__*След съгласие чрез бутон "Ок" ще се генерира файл, който е готов за анализ и изготвяне на седмичната справка.*__
> <img src="./images/seccondPage5.JPG">
> <br /><br />
> <img src="./images/seccondPage6.JPG">
> <br /><br />
> <img src="./images/seccondPage7.JPG">
> </br>
> 
</br>

<table  border = '0' style="width: 100%;">
  <tr>
    <td>
        <img align="left" width="116" height="116" src="./favicon.ico" >
    </td>
    <td> 

## CSICorp INTELLIGENT SYSTEMS

</tr>
</table>
<br />

#### ***Адрес за тест на приложението ***[CSICorp](http://asarel.csicorp.eu)***. Файлове с които да тествате изчислението на данни и генериране на xlsx файл - [вариан 1](./DataForTest/9%20week.zip), [вариант 2](./DataForTest/8%20week.zip)***