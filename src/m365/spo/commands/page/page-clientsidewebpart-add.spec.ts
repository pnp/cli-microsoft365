import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./page-clientsidewebpart-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.PAGE_CLIENTSIDEWEBPART_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  const clientSideWebParts = {
    value: [
      {
        ComponentType: 1,
        Id: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
        Manifest:
          '{"preconfiguredEntries":[{"title":{"default":"Bing maps","ar-SA":"خرائط Bing","az-Latn-AZ":"Bing xəritələr","bg-BG":"Карти на Bing","bs-Latn-BA":"Bing karte","ca-ES":"Mapes del Bing","cs-CZ":"Mapy Bing","cy-GB":"Mapiau Bing","da-DK":"Bing Kort","de-DE":"Bing Karten","el-GR":"Χάρτες Bing","en-US":"Bing maps","es-ES":"Mapas de Bing","et-EE":"Bingi kaardid","eu-ES":"Bing Mapak","fi-FI":"Bing-kartat","fr-FR":"Bing Cartes","ga-IE":"Mapaí Bing","gl-ES":"Mapas de Bing","he-IL":"מפות Bing","hi-IN":"Bing मानचित्र","hr-HR":"Bing karte","hu-HU":"Bing Térképek","id-ID":"Peta Bing","it-IT":"Bing Maps","ja-JP":"Bing 地図","kk-KZ":"Bing карталары","ko-KR":"Bing 지도","lt-LT":"„Bing“ žemėlapiai","lv-LV":"Bing kartes","mk-MK":"Мапи Bing","ms-MY":"Peta Bing","nb-NO":"Bing-kart","nl-NL":"Bing Kaarten","pl-PL":"Mapy Bing","prs-AF":"نقشه های Bing","pt-BR":"Bing Mapas","pt-PT":"Mapas Bing","qps-ploc":"[!!--##Ƀĭńġ màƿš##--!!]","qps-ploca":"Ɓįƞġ mȃƿš","ro-RO":"Hărți Bing","ru-RU":"Карты Bing","sk-SK":"Mapy Bing","sl-SI":"Zemljevidi Bing","sr-Cyrl-RS":"Bing мапе","sr-Latn-RS":"Bing mape","sv-SE":"Bing-kartor","th-TH":"Bing Maps","tr-TR":"Bing haritalar","uk-UA":"Карти Bing","vi-VN":"Bản đồ Bing","zh-CN":"必应地图","zh-TW":"Bing 地圖服務","en-GB":"Bing maps","en-NZ":"Bing maps","en-IE":"Bing maps","en-AU":"Bing maps"},"description":{"default":"Display a key location on a map","ar-SA":"عرض موقع رئيسي على خريطة","az-Latn-AZ":"Xəritədə əsas yeri göstərin","bg-BG":"Показване на ключово местоположение на картата","bs-Latn-BA":"Prikažite ključnu lokaciju na karti","ca-ES":"Mostreu una ubicació clau en un mapa","cs-CZ":"Umožňuje zobrazit důležitou polohu na mapě.","cy-GB":"Dangos lleoliad allweddol ar y map","da-DK":"Vis en nøgleplacering på et kort","de-DE":"Einen wichtigen Ort auf einer Karte anzeigen.","el-GR":"Εμφανίστε μια σημαντική τοποθεσία σε έναν χάρτη","en-US":"Display a key location on a map","es-ES":"Muestra la ubicación clave en un mapa.","et-EE":"Saate kaardil kuvada põhiasukoha","eu-ES":"Erakutsi kokaleku nagusi bat mapa batean","fi-FI":"Näytä merkittävä sijainti kartalla","fr-FR":"Affichez un emplacement clé sur une carte","ga-IE":"Taispeáin príomhshuíomh ar mhapa","gl-ES":"Fixa unha localización importante no mapa.","he-IL":"הצג מיקום מרכזי במפה","hi-IN":"मानचित्र पर कोई मुख्य स्थान प्रदर्शित करें","hr-HR":"Prikaz ključnog mjesta na karti","hu-HU":"Elsődleges hely megjelenítése a térképen","id-ID":"Tampilkan lokasi penting di peta","it-IT":"Visualizzare una posizione chiave su una mappa","ja-JP":"主要な場所をマップに表示します","kk-KZ":"Негізгі орынды картадан көрсету","ko-KR":"지도에서 핵심 위치를 표시합니다.","lt-LT":"Rodyti pagrindinę vietą žemėlapyje","lv-LV":"Parādīt svarīgu atrašanās vietu kartē","mk-MK":"Покажете клучна локација на мапата","ms-MY":"Papar lokasi utama pada peta","nb-NO":"Vis en viktig plassering på et kart","nl-NL":"Belangrijke locatie op een kaart weergeven","pl-PL":"Wyświetl kluczową lokalizację na mapie","prs-AF":"نمایش یک موقعیت کلیدی در نقشه","pt-BR":"Exiba um local importante em um mapa","pt-PT":"Apresentar uma localização importante num mapa","qps-ploc":"[!!--##Ɖıŝƿľāƴ ä ķĕɏ ļơćąţıőň ŏń ȃ māƥ##--!!]","qps-ploca":"Ɗīšƿłâŷ á ķȇɏ ŀőċáťĩōņ ŏň ā màƥ","ro-RO":"Afișați o locație cheie pe o hartă","ru-RU":"Отображение ключевого расположения на карте","sk-SK":"Umožňuje zobraziť na mape dôležitú polohu","sl-SI":"Prikaži ključno lokacijo na zemljevidu.","sr-Cyrl-RS":"Прикажите кључну локацију на мапи","sr-Latn-RS":"Prikažite ključnu lokaciju na mapi","sv-SE":"Visa en viktig plats på en karta","th-TH":"แสดงตำแหน่งที่ตั้งหลักบนแผนที่","tr-TR":"Önemli bir konumu haritada görüntüleyin","uk-UA":"Відображення ключового розташування на карті","vi-VN":"Hiển thị vị trí quan trọng trên bản đồ","zh-CN":"在地图上显示关键位置","zh-TW":"在地圖上顯示重要位置","en-GB":"Display a key location on a map","en-NZ":"Display a key location on a map","en-IE":"Display a key location on a map","en-AU":"Display a key location on a map"},"officeFabricIconFontName":"MapPin","iconImageUrl":null,"groupId":"cf066440-0614-43d6-98ae-0b31cf14c7c3","group":{"default":"Media and Content","ar-SA":"الوسائط والمحتويات","az-Latn-AZ":"Media və Məzmun","bg-BG":"Мултимедия и съдържание","bs-Latn-BA":"Mediji i sadržaj","ca-ES":"Fitxers multimèdia i contingut","cs-CZ":"Multimédia a obsah","cy-GB":"Cyfryngau a Chynnwys","da-DK":"Medier og indhold","de-DE":"Medien und Inhalt","el-GR":"Πολυμέσα και περιεχόμενο","en-US":"Media and Content","es-ES":"Contenido y elementos multimedia","et-EE":"Meediumid ja sisu","eu-ES":"Multimedia eta edukia","fi-FI":"Media ja sisältö","fr-FR":"Média et contenu","ga-IE":"Meáin agus inneachar","gl-ES":"Contido e elementos multimedia","he-IL":"מדיה ותוכן","hi-IN":"मीडिया और सामग्री","hr-HR":"Mediji i sadržaj","hu-HU":"Média és tartalom","id-ID":"Media dan Konten","it-IT":"Elementi multimediali e contenuto","ja-JP":"メディアおよびコンテンツ","kk-KZ":"Мультимедиа және контент","ko-KR":"미디어 및 콘텐츠","lt-LT":"Medija ir turinys","lv-LV":"Multivide un saturs","mk-MK":"Медиуми и содржина","ms-MY":"Media dan Kandungan","nb-NO":"Medier og innhold","nl-NL":"Media en inhoud","pl-PL":"Multimedia i zawartość","prs-AF":"مطبوعات و محتوا","pt-BR":"Mídia e Conteúdo","pt-PT":"Multimédia e Conteúdo","qps-ploc":"[!!--##Mēđĩǻ ȃņđ Ĉōńťȇņť##--!!]","qps-ploca":"Měďīǻ ǻńď Ċōƞţȅƞţ","ro-RO":"Media și conținut","ru-RU":"Мультимедиа и контент","sk-SK":"Médiá a obsah","sl-SI":"Predstavnost in vsebina","sr-Cyrl-RS":"Медији и садржај","sr-Latn-RS":"Mediji i sadržaj","sv-SE":"Media och innehåll","th-TH":"สื่อและเนื้อหา","tr-TR":"Medya ve İçerik","uk-UA":"Мультимедіа та вміст","vi-VN":"Phương tiện và Nội dung","zh-CN":"媒体和内容","zh-TW":"媒體及內容","en-GB":"Media and Content","en-NZ":"Media and Content","en-IE":"Media and Content","en-AU":"Media and Content"},"properties":{"pushPins":[],"maxNumberOfPushPins":1,"shouldShowPushPinTitle":true,"zoomLevel":12,"mapType":"road"}}],"disabledOnClassicSharepoint":false,"searchablePropertyNames":null,"linkPropertyNames":null,"imageLinkPropertyNames":null,"hiddenFromToolbox":false,"supportsFullBleed":false,"requiredCapabilities":{"BingMapsKey":true},"isolationLevel":"None","version":"1.2.0","alias":"BingMapWebPart","preloadComponents":null,"isInternal":true,"loaderConfig":{"internalModuleBaseUrls":["https://spoprod-a.akamaihd.net/files/"],"entryModuleId":"sp-bing-map-webpart-bundle","exportName":null,"scriptResources":{"sp-bing-map-webpart-bundle":{"type":"localizedPath","shouldNotPreload":false,"id":null,"version":null,"failoverPath":null,"path":null,"defaultPath":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_default_f9778f1c0ab9952745886d1d605776ef.js","paths":{"ar-SA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","ar":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","tzm-Latn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","ku":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","syr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","az-Latn-AZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_az-latn-az_85fe3797228bbb31c90827117a1e88b6.js","bg-BG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_bg-bg_c18a087f3cd05dfb53d2a764c84ee8e5.js","bs-Latn-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_bs-latn-ba_6d0592293c02851de2b8d6426de41275.js","ca-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ca-es_67b4773f20c230a76857bd60bf748dac.js","cs-CZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cs-cz_64b585926105d1854c19a591a96d4742.js","cy-GB":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cy-gb_d65e5b362d048891264360153e8b3925.js","cy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cy-gb_d65e5b362d048891264360153e8b3925.js","da-DK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","fo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","kl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","de-DE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","de":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","dsb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","rm":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","hsb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","el-GR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_el-gr_fb1d150ca6185cc1aeb3077b88b19884.js","el":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_el-gr_fb1d150ca6185cc1aeb3077b88b19884.js","en-US":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","bn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","chr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","dv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","div":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","en":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","fil":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","haw":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","iu":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","lo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","moh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","es-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","gn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","quz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","es":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","ca-ES-valencia":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","et-EE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_et-ee_4843ca005a2b35f4d37641aa010c9841.js","eu-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_eu-es_9a9719aea4f15cae349923e677971eb9.js","fi-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","sms":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","se-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","se-Latn-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","fr-FR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","gsw":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","br":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","tzm-Tfng":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","co":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","fr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","ff":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","lb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","mg":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","oc":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","zgh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","wo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","ga-IE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ga-ie_293592b7ef4495b6f2ac24ee7209a4ba.js","gl-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_gl-es_73fd41fe859fc930e937f011dd817df0.js","he-IL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_he-il_99de68842371a4fc44f2c0b6ff9a5728.js","hi-IN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hi-in_4b5ca5e56fe8161732825c681204a7b8.js","hi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hi-in_4b5ca5e56fe8161732825c681204a7b8.js","hr-HR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hr-hr_f390168a1e0ecbe1449b3671bb8d40e4.js","hu-HU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hu-hu_45735154b266d9a5c4c22d5f58ff357b.js","id-ID":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_id-id_16c289d93cc0beeff4bc67251a3eee1f.js","jv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_id-id_16c289d93cc0beeff4bc67251a3eee1f.js","it-IT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_it-it_01d4cf9abd13c78a05176b5603fcd910.js","it":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_it-it_01d4cf9abd13c78a05176b5603fcd910.js","ja-JP":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ja-jp_1e2c2f23c172b5b6cb58db6a93c22377.js","kk-KZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_kk-kz_b964c6292da4da296886bdb4333612af.js","ko-KR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ko-kr_732fca9752fe7e578a07302ec1664b60.js","lt-LT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_lt-lt_bf02cafe5c3eb67d6c226b57dec12e97.js","lv-LV":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_lv-lv_b1663a1f172b8fc522347238c799f7fd.js","mk-MK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_mk-mk_e98c9da753231452008db194cb735a2d.js","ms-MY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ms-my_2f6050a7c3ce3b6ffbb47689a4abaf29.js","ms":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ms-my_2f6050a7c3ce3b6ffbb47689a4abaf29.js","nb-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","no":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","smj-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","smj-Latn-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","se-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","se-Latn-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","sma-Latn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","sma-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nl-NL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","nl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","fy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","pl-PL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pl-pl_605a6ac656d59c71bfd6db4678c92ff2.js","prs-AF":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","gbz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","ps":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","pt-BR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-br_9008a60143e5a19971076daced802a6e.js","pt-PT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-pt_224b200ba7df98bc7e2dfdad6762b7d9.js","pt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-pt_224b200ba7df98bc7e2dfdad6762b7d9.js","qps-ploc":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_qps-ploc_7cc2e80271e05822b85637079fbbc8e6.js","qps-ploca":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_qps-ploca_cc832fbdf67d10b7072faeaecf7a48f7.js","ro-RO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ro-ro_4d43cca645941d769735433b76b88c92.js","ro":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ro-ro_4d43cca645941d769735433b76b88c92.js","ru-RU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ru":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ba":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","be":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ky":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","mn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","sah":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tg":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","sk-SK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sk-sk_43b75662b363ee2797c81026449182af.js","sk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sk-sk_43b75662b363ee2797c81026449182af.js","sl-SI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sl-si_4fad5739556858cca3f39917cb8b0cbd.js","sl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sl-si_4fad5739556858cca3f39917cb8b0cbd.js","sr-Cyrl-RS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-cyrl-rs_c2b5a41feb0599111c17612d389880cb.js","sr-Cyrl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-cyrl-rs_c2b5a41feb0599111c17612d389880cb.js","sr-Latn-RS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-latn-rs_8ccc6f642938bf9e96f231e693a28fd2.js","sr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-latn-rs_8ccc6f642938bf9e96f231e693a28fd2.js","sv-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","smj":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","se":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sma-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sma-Latn-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","th-TH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_th-th_04b7da4b93097cb2e7f295e1afd80d19.js","th":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_th-th_04b7da4b93097cb2e7f295e1afd80d19.js","tr-TR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_tr-tr_d81724d2147360856bba8c7de5acc476.js","tr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_tr-tr_d81724d2147360856bba8c7de5acc476.js","uk-UA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_uk-ua_9365a5370d64d5650444397ab89e4b5c.js","vi-VN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_vi-vn_e45cb5b2cb1a7bcb517fb17e04188820.js","vi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_vi-vn_e45cb5b2cb1a7bcb517fb17e04188820.js","zh-CN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","zh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","mn-Mong":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","bo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","ug":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","ii":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","zh-TW":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-HK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-CHT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-Hant":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-MO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","en-GB":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sq":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","am":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","hy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","mk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","bs":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","my":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","dz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-CY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-EG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-IL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-IS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-JO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-KE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-KW":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-MK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-MT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-PK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-QA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-SA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-LK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-AE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-VN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","is":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","km":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","kh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","mt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","fa":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","gd":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sr-Cyrl-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sr-Latn-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sd":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","si":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","so":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","ti-ET":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","uz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-NZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-nz_f9778f1c0ab9952745886d1d605776ef.js","en-IE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-ie_f9778f1c0ab9952745886d1d605776ef.js","en-AU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-SG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-HK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-MY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-PH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-TT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-AZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-BH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-BN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-ID":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","mi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js"},"globalName":null,"globalDependencies":null},"react":{"type":"component","shouldNotPreload":false,"id":"0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d","version":"15.6.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/office-ui-fabric-react-bundle":{"type":"component","shouldNotPreload":false,"id":"02a01e42-69ab-403d-8a16-acd128661f8e","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/load-themed-styles":{"type":"component","shouldNotPreload":false,"id":"229b8d08-79f3-438b-8c21-4613fc877abd","version":"0.1.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/odsp-utilities-bundle":{"type":"component","shouldNotPreload":false,"id":"cc2cc925-b5be-41bb-880a-f0f8030c6aff","version":"4.1.3","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-telemetry":{"type":"component","shouldNotPreload":false,"id":"8217e442-8ed3-41fd-957d-b112e841286a","version":"0.2.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-core-library":{"type":"component","shouldNotPreload":false,"id":"7263c7d0-1d6a-45ec-8d85-d4d1d234171b","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-webpart-base":{"type":"component","shouldNotPreload":false,"id":"974a7777-0990-4136-8fa6-95d80114c2e0","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-webpart-shared":{"type":"component","shouldNotPreload":false,"id":"914330ee-2df2-4f6e-a858-30c23a812408","version":"0.1.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-component-utilities":{"type":"component","shouldNotPreload":false,"id":"8494e7d7-6b99-47b2-a741-59873e42f16f","version":"0.2.1","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"react-dom":{"type":"component","shouldNotPreload":false,"id":"aa0a46ec-1505-43cd-a44a-93f3a5aa460a","version":"15.6.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-lodash-subset":{"type":"component","shouldNotPreload":false,"id":"73e1dc6c-8441-42cc-ad47-4bd3659f8a3a","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-diagnostics":{"type":"component","shouldNotPreload":false,"id":"78359e4b-07c2-43c6-8d0b-d060b4d577e8","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-bingmap":{"type":"component","shouldNotPreload":false,"id":"ab22169a-a644-4b69-99a2-4295eb0f633c","version":"0.0.1","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null}}},"manifestVersion":2,"id":"e377ea37-9047-43b9-8cdb-a761be2f8e09","componentType":"WebPart"}',
        ManifestType: 1,
        Name: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
        Status: 0
      }
    ]
  };
  // function to replace the dynamically generated web part instance id
  // with a static GUID
  const replaceId = (str: string): string => {
    const id: string[] = [];
    let s = str.replace(/(\"instanceId\\\"\:\\\")([^\\]+)(\\\")/g, (match: string, p1: string, p2: string, p3: string, offset: string, fullString: string): string => {
      id.push(p2);
      return `${p1}89c644b3-f69c-4e84-85d7-dfa04c6163b5${p3}`;
    });
    id.forEach(idx => {
      s = s.replace(`\\"id\\":\\"${idx}\\"`, `\\"id\\":\\"89c644b3-f69c-4e84-85d7-dfa04c6163b5\\"`);
    });

    return s;
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PAGE_CLIENTSIDEWEBPART_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('checks out page if not checked out by the current user', (done) => {
    let checkedOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": false,
          "CanvasContent1": null
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkedOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
      }
    }, () => {
      try {
        assert.deepEqual(checkedOut, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('checks out page if not checked out by the current user (debug)', (done) => {
    let checkedOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": false,
          "CanvasContent1": null
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkedOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
      }
    }, () => {
      try {
        assert.deepEqual(checkedOut, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t check out page if checked out by the current user', (done) => {
    let checkingOut = false;
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": null
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/checkoutpage`) > -1) {
        checkingOut = true;
        return Promise.resolve({});
      }

      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/home.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        pageName: 'home',
        webUrl: 'https://contoso.sharepoint.com/sites/newsletter',
        webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
      }
    }, () => {
      try {
        assert.deepEqual(checkingOut, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds web part to an empty column when no order specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "layoutIndex": 1,
                  "controlIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part to an empty column when order 1 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 1
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "layoutIndex": 1,
                  "controlIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part to an empty column when order 5 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 5
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "layoutIndex": 1,
                  "controlIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the end of the column with one web part when no order specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the beginning of the column with one web part when order 1 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 1
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the end of the column with one web part when order 2 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 2
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the end of the column with multiple web part when no order specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"controlType\":3,\"displayMode\":2,\"id\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1.5,\"layoutIndex\":1,\"sectionFactor\":8},\"webPartId\":\"62cac389-787f-495d-beca-e11786162ef4\",\"emphasis\":{},\"reservedHeight\":321,\"reservedWidth\":757,\"webPartData\":{\"id\":\"62cac389-787f-495d-beca-e11786162ef4\",\"instanceId\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"title\":\"Countdown Timer\",\"description\":\"This web part is used to allow a site admin to count down/up to an important event.\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{\"buttonURL\":null}},\"dataVersion\":\"2.1\",\"properties\":{\"showButton\":false,\"countDate\":\"Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)\",\"title\":\"\",\"description\":\"\",\"countDirection\":\"COUNT_DOWN\",\"dateDisplay\":\"DAY_HOUR_MINUTE_SECOND\",\"buttonText\":\"\"}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "controlIndex": 2,
                  "layoutIndex": 1,
                  "sectionFactor": 8
                },
                "webPartId": "62cac389-787f-495d-beca-e11786162ef4",
                "emphasis": {},
                "reservedHeight": 321,
                "reservedWidth": 757,
                "webPartData": {
                  "id": "62cac389-787f-495d-beca-e11786162ef4",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Countdown Timer",
                  "description": "This web part is used to allow a site admin to count down/up to an important event.",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {
                      "buttonURL": null
                    }
                  },
                  "dataVersion": "2.1",
                  "properties": {
                    "showButton": false,
                    "countDate": "Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)",
                    "title": "",
                    "description": "",
                    "countDirection": "COUNT_DOWN",
                    "dateDisplay": "DAY_HOUR_MINUTE_SECOND",
                    "buttonText": ""
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 3,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the beginning of the column with multiple web part when order 1 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"controlType\":3,\"displayMode\":2,\"id\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1.5,\"layoutIndex\":1,\"sectionFactor\":8},\"webPartId\":\"62cac389-787f-495d-beca-e11786162ef4\",\"emphasis\":{},\"reservedHeight\":321,\"reservedWidth\":757,\"webPartData\":{\"id\":\"62cac389-787f-495d-beca-e11786162ef4\",\"instanceId\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"title\":\"Countdown Timer\",\"description\":\"This web part is used to allow a site admin to count down/up to an important event.\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{\"buttonURL\":null}},\"dataVersion\":\"2.1\",\"properties\":{\"showButton\":false,\"countDate\":\"Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)\",\"title\":\"\",\"description\":\"\",\"countDirection\":\"COUNT_DOWN\",\"dateDisplay\":\"DAY_HOUR_MINUTE_SECOND\",\"buttonText\":\"\"}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 1
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "controlIndex": 3,
                  "layoutIndex": 1,
                  "sectionFactor": 8
                },
                "webPartId": "62cac389-787f-495d-beca-e11786162ef4",
                "emphasis": {},
                "reservedHeight": 321,
                "reservedWidth": 757,
                "webPartData": {
                  "id": "62cac389-787f-495d-beca-e11786162ef4",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Countdown Timer",
                  "description": "This web part is used to allow a site admin to count down/up to an important event.",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {
                      "buttonURL": null
                    }
                  },
                  "dataVersion": "2.1",
                  "properties": {
                    "showButton": false,
                    "countDate": "Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)",
                    "title": "",
                    "description": "",
                    "countDirection": "COUNT_DOWN",
                    "dateDisplay": "DAY_HOUR_MINUTE_SECOND",
                    "buttonText": ""
                  }
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part in the middle of the column with multiple web part when order 2 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"controlType\":3,\"displayMode\":2,\"id\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1.5,\"layoutIndex\":1,\"sectionFactor\":8},\"webPartId\":\"62cac389-787f-495d-beca-e11786162ef4\",\"emphasis\":{},\"reservedHeight\":321,\"reservedWidth\":757,\"webPartData\":{\"id\":\"62cac389-787f-495d-beca-e11786162ef4\",\"instanceId\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"title\":\"Countdown Timer\",\"description\":\"This web part is used to allow a site admin to count down/up to an important event.\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{\"buttonURL\":null}},\"dataVersion\":\"2.1\",\"properties\":{\"showButton\":false,\"countDate\":\"Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)\",\"title\":\"\",\"description\":\"\",\"countDirection\":\"COUNT_DOWN\",\"dateDisplay\":\"DAY_HOUR_MINUTE_SECOND\",\"buttonText\":\"\"}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 2
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "controlIndex": 3,
                  "layoutIndex": 1,
                  "sectionFactor": 8
                },
                "webPartId": "62cac389-787f-495d-beca-e11786162ef4",
                "emphasis": {},
                "reservedHeight": 321,
                "reservedWidth": 757,
                "webPartData": {
                  "id": "62cac389-787f-495d-beca-e11786162ef4",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Countdown Timer",
                  "description": "This web part is used to allow a site admin to count down/up to an important event.",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {
                      "buttonURL": null
                    }
                  },
                  "dataVersion": "2.1",
                  "properties": {
                    "showButton": false,
                    "countDate": "Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)",
                    "title": "",
                    "description": "",
                    "countDirection": "COUNT_DOWN",
                    "dateDisplay": "DAY_HOUR_MINUTE_SECOND",
                    "buttonText": ""
                  }
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds web part at the end of the column with multiple web part when order 5 specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"controlType\":3,\"displayMode\":2,\"id\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"controlIndex\":1.5,\"layoutIndex\":1,\"sectionFactor\":8},\"webPartId\":\"62cac389-787f-495d-beca-e11786162ef4\",\"emphasis\":{},\"reservedHeight\":321,\"reservedWidth\":757,\"webPartData\":{\"id\":\"62cac389-787f-495d-beca-e11786162ef4\",\"instanceId\":\"230b9699-d4ed-414b-8a83-9b251297c384\",\"title\":\"Countdown Timer\",\"description\":\"This web part is used to allow a site admin to count down/up to an important event.\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{\"buttonURL\":null}},\"dataVersion\":\"2.1\",\"properties\":{\"showButton\":false,\"countDate\":\"Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)\",\"title\":\"\",\"description\":\"\",\"countDirection\":\"COUNT_DOWN\",\"dateDisplay\":\"DAY_HOUR_MINUTE_SECOND\",\"buttonText\":\"\"}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          order: 5
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "controlIndex": 2,
                  "layoutIndex": 1,
                  "sectionFactor": 8
                },
                "webPartId": "62cac389-787f-495d-beca-e11786162ef4",
                "emphasis": {},
                "reservedHeight": 321,
                "reservedWidth": 757,
                "webPartData": {
                  "id": "62cac389-787f-495d-beca-e11786162ef4",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Countdown Timer",
                  "description": "This web part is used to allow a site admin to count down/up to an important event.",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {
                      "buttonURL": null
                    }
                  },
                  "dataVersion": "2.1",
                  "properties": {
                    "showButton": false,
                    "countDate": "Sun Apr 07 2019 22:00:00 GMT+0200 (Central European Summer Time)",
                    "title": "",
                    "description": "",
                    "countDirection": "COUNT_DOWN",
                    "dateDisplay": "DAY_HOUR_MINUTE_SECOND",
                    "buttonText": ""
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 3,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a standard web part at the end of the column with one web part when no order specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a standard web part in a default section when no section exists on page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "zoneIndex": 1,
                  "sectionFactor": 12,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a standard web part with properties at the end of the column with one web part when no order specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap',
          webPartProperties: '{"title":"Location"}'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road",
                    "title": "Location"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error when adding Client Side Web Part to non-existing page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/foo.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'The file /sites/team-a/sitepages/foo.aspx does not exist' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'foo.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file /sites/team-a/sitepages/foo.aspx does not exist')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error if WebPart API does not respond properly', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles OData error when adding Client Side Web Part to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles WebPart properties error when adding Client Side Web Part to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          webPartProperties: '{"foo", }'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles invalid specified WebPart Id error when adding Client Side Web Part to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-aaaaaaaaaa'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`There is no available WebPart with Id e377ea37-9047-43b9-8cdb-aaaaaaaaaa.`)));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles error if target page is not a modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.reject({
          error: {
            "odata.error": {
              "code": "-1, Microsoft.SharePoint.Client.ClientServiceException",
              "message": {
                "lang": "en-US",
                "value": "This page does not have the site page content type. Only site pages can be served with this API."
              }
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-aaaaaaaaaa'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`This page does not have the site page content type. Only site pages can be served with this API.`)));
          done();
        }
        catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles invalid section error when adding Client Side Web Part to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          section: 8
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Invalid section '8'")));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles invalid column error when adding Client Side Web Part to modern page', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
          section: 1,
          column: 7
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Invalid column '7'")));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a web part using web part data', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap',
          webPartData: '{"id": "e377ea37-9047-43b9-8cdb-a761be2f8e09","instanceId": "f2f0ee32-eba5-47f9-9aa1-24f99661ecd1","title": "Bing Maps","description": "Display a key location on a map","serverProcessedContent": {"htmlStrings": {},"searchablePlainTexts": {},"imageSources": {},"links": {}},"properties": {"pushPins": [{"location": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1,"name": "Microsoft Way, Redmond, WA 98052"},"defaultTitle": "One Microsoft Way, Building 32","defaultAddress": "Microsoft Way, Redmond, WA 98052","title": "One Microsoft Way, Building 32","address": "Microsoft Way, Redmond, WA 98052"}],"maxNumberOfPushPins": 1,"shouldShowPushPinTitle": true,"zoomLevel": 12,"mapType": "road","center": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1}}}'
        }
      },
      () => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {"pushPins": [{"location": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1,"name": "Microsoft Way, Redmond, WA 98052"},"defaultTitle": "One Microsoft Way, Building 32","defaultAddress": "Microsoft Way, Redmond, WA 98052","title": "One Microsoft Way, Building 32","address": "Microsoft Way, Redmond, WA 98052"}],"maxNumberOfPushPins": 1,"shouldShowPushPinTitle": true,"zoomLevel": 12,"mapType": "road","center": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1}},
                  "title": "Bing Maps",
                  "serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a web part using web part data (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap',
          webPartData: '{"id": "e377ea37-9047-43b9-8cdb-a761be2f8e09","instanceId": "f2f0ee32-eba5-47f9-9aa1-24f99661ecd1","title": "Bing Maps","description": "Display a key location on a map","serverProcessedContent": {"htmlStrings": {},"searchablePlainTexts": {},"imageSources": {},"links": {}},"properties": {"pushPins": [{"location": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1,"name": "Microsoft Way, Redmond, WA 98052"},"defaultTitle": "One Microsoft Way, Building 32","defaultAddress": "Microsoft Way, Redmond, WA 98052","title": "One Microsoft Way, Building 32","address": "Microsoft Way, Redmond, WA 98052"}],"maxNumberOfPushPins": 1,"shouldShowPushPinTitle": true,"zoomLevel": 12,"mapType": "road","center": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1}}}'
        }
      },
      () => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {"pushPins": [{"location": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1,"name": "Microsoft Way, Redmond, WA 98052"},"defaultTitle": "One Microsoft Way, Building 32","defaultAddress": "Microsoft Way, Redmond, WA 98052","title": "One Microsoft Way, Building 32","address": "Microsoft Way, Redmond, WA 98052"}],"maxNumberOfPushPins": 1,"shouldShowPushPinTitle": true,"zoomLevel": 12,"mapType": "road","center": {"latitude": 47.6405250000244,"longitude": -122.129415000122,"altitude": 0,"altitudeReference": -1}},
                  "title": "Bing Maps",
                  "serverProcessedContent":{"htmlStrings":{},"searchablePlainTexts":{},"imageSources":{},"links":{}}
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('adds a web part with dynamicDataPaths and dynamicDataValues', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"controlType\":3,\"displayMode\":2,\"id\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"position\":{\"controlIndex\":0.5,\"sectionIndex\":1,\"sectionFactor\":8,\"zoneIndex\":1,\"layoutIndex\":1},\"webPartId\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"emphasis\":{},\"reservedHeight\":127,\"reservedWidth\":757,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823\",\"instanceId\":\"4dae46b3-b059-4930-8495-0920cec4faa0\",\"title\":\"Weather\",\"description\":\"Show current weather conditions on your page\",\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"temperatureUnit\":\"F\",\"locations\":[{\"latitude\":47.604,\"longitude\":-122.329,\"name\":\"Seattle, United States\",\"showCustomizedDisplayName\":false}]}}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action(
      {
        options: {
          debug: true,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          standardWebPart: 'BingMap',
          webPartData: '{"id": "e377ea37-9047-43b9-8cdb-a761be2f8e09","instanceId": "f2f0ee32-eba5-47f9-9aa1-24f99661ecd1","title": "Bing Maps","description": "Display a key location on a map","dataVersion": "1.0", "dynamicDataPaths":{"dynamicProperty0":"WebPart.2bacb933-9f9d-457f-bfa5-b00bfc9cd625.69800bc3-0d7c-495c-a5b6-3423f226d5c5:queryText"},"dynamicDataValues":{"dynamicProperty1":""}}'
        }
      },
      () => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                "emphasis": {},
                "reservedHeight": 127,
                "reservedWidth": 757,
                "addedFromPersistedData": true,
                "webPartData": {
                  "id": "868ac3c3-cad7-4bd6-9a1c-14dc5cc8e823",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "title": "Weather",
                  "description": "Show current weather conditions on your page",
                  "serverProcessedContent": {
                    "htmlStrings": {},
                    "searchablePlainTexts": {},
                    "imageSources": {},
                    "links": {}
                  },
                  "dataVersion": "1.2",
                  "properties": {
                    "temperatureUnit": "F",
                    "locations": [
                      {
                        "latitude": 47.604,
                        "longitude": -122.329,
                        "name": "Seattle, United States",
                        "showCustomizedDisplayName": false
                      }
                    ]
                  }
                }
              },
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "controlIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "zoneIndex": 1,
                  "layoutIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing Maps",
                  "dynamicDataPaths":{"dynamicProperty0":"WebPart.2bacb933-9f9d-457f-bfa5-b00bfc9cd625.69800bc3-0d7c-495c-a5b6-3423f226d5c5:queryText"},"dynamicDataValues":{"dynamicProperty1":""}
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('correctly handles sections in reverse order', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')?$select=CanvasContent1,IsPageCheckedOutToCurrentUser`) > -1) {
        return Promise.resolve({
          "IsPageCheckedOutToCurrentUser": true,
          "CanvasContent1": "[{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":1,\"sectionFactor\":8,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":2,\"sectionIndex\":2,\"sectionFactor\":4,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":1,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"displayMode\":2,\"position\":{\"zoneIndex\":1,\"sectionIndex\":2,\"sectionFactor\":6,\"layoutIndex\":1},\"emphasis\":{}},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isDefaultThumbnail\":true}}]"
        });
      }

      if ((opts.url as string).indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
        return Promise.resolve(clientSideWebParts);
      }

      return Promise.reject('Invalid request');
    });

    let body: string = '';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/sitepages/pages/GetByUrl('sitepages/page.aspx')/savepage`) > -1) {
        body = opts.body;
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    
    cmdInstance.action(
      {
        options: {
          debug: false,
          pageName: 'page.aspx',
          webUrl: 'https://contoso.sharepoint.com/sites/team-a',
          webPartId: 'e377ea37-9047-43b9-8cdb-a761be2f8e09'
        }
      },
      (err?: any) => {
        try {
          assert.strictEqual(replaceId(JSON.stringify(body)), JSON.stringify({
            CanvasContent1: JSON.stringify([
              {
                "controlType": 3,
                "displayMode": 2,
                "id": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                "position": {
                  "zoneIndex": 2,
                  "sectionIndex": 1,
                  "sectionFactor": 8,
                  "layoutIndex": 1,
                  "controlIndex": 1
                },
                "webPartId": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                "emphasis": {},
                "webPartData": {
                  "dataVersion": "1.0",
                  "description": "Display a key location on a map",
                  "id": "e377ea37-9047-43b9-8cdb-a761be2f8e09",
                  "instanceId": "89c644b3-f69c-4e84-85d7-dfa04c6163b5",
                  "properties": {
                    "pushPins": [],
                    "maxNumberOfPushPins": 1,
                    "shouldShowPushPinTitle": true,
                    "zoomLevel": 12,
                    "mapType": "road"
                  },
                  "title": "Bing maps"
                }
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 2,
                  "sectionIndex": 2,
                  "sectionFactor": 4,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 1,
                  "sectionFactor": 6,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "displayMode": 2,
                "position": {
                  "zoneIndex": 1,
                  "sectionIndex": 2,
                  "sectionFactor": 6,
                  "layoutIndex": 1
                },
                "emphasis": {}
              },
              {
                "controlType": 0,
                "pageSettingsSlice": {
                  "isDefaultDescription": true,
                  "isDefaultThumbnail": true
                }
              }
            ])
          }));
          done();
        } catch (e) {
          done(e);
        }
      }
    );
  });

  it('supports debug mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports verbose mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying page name', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--pageName') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webPartId', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webPartId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying standardWebPart', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--standardWebPart') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webPartProperties', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--webPartProperties') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying section', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--section') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying column', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--column') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying order', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option.indexOf('--order') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if webUrl is not an absolute URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'foo', webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'http://foo',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if either webPartId or standardWebPart parameters are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webPartId and standardWebPart parameters are both specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        standardWebPart: 'BingMap'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webPartId value is not valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: 'FooBar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webPartProperties and webPartData are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        webPartProperties: '{}',
        webPartData: '{}'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webPartProperties value is not valid JSON', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        webPartProperties: '{Foo:bar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webPartProperties value is valid JSON', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        webPartProperties: '{}'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if webPartData value is not valid JSON', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        webPartData: '{Foo:bar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webPartData value is valid JSON', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        webPartData: '{}'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if standardWebPart is not valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', standardWebPart: 'Foo' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when name and webURL specified, webUrl is a valid SharePoint URL and webPartId is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name and webURL specified, webUrl is a valid SharePoint URL and standardWebPart is specified instead of webPartId', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', standardWebPart: 'BingMap' }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when name has no extension', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if standardWebPart is valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { pageName: 'page.aspx', webUrl: 'https://contoso.sharepoint.com', standardWebPart: 'BingMap' }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if section has invalid (negative) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        section: -1
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if section has invalid (non number) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        section: 'foobar'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if column has invalid (negative) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        column: -1
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if column has invalid (non number) value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        pageName: 'page.aspx',
        webUrl: 'https://contoso.sharepoint.com',
        webPartId: '3ede60d3-dc2c-438b-b5bf-cc40bb2351e1',
        column: 'foobar'
      }
    });
    assert.notStrictEqual(actual, true);
  });
});
