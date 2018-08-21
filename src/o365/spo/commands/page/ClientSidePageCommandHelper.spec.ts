import * as sinon from 'sinon';
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { ClientSidePage, ClientSideWebpart, ClientSideControlData, ClientSideText } from './clientsidepages';
import ClientSidePageCommandHelper, { ICommandContext } from './ClientSidePageCommandHelper';
import { StandardWebPartUtils } from '../../common/StandardWebPartTypes';

describe('ClientSidePageCommandHelper', () => {
	let log: string[];
  let commandContext: ICommandContext;
  let loggerSpy: sinon.SinonSpy;

	const bingMapWebPartId = 'e377ea37-9047-43b9-8cdb-a761be2f8e09';
	const bingMapWebPartName = 'BingMap';
	const bingMapWebPartDefinition = {
		ComponentType: 1,
		Id: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
		Manifest:
			'{"preconfiguredEntries":[{"title":{"default":"Bing maps","ar-SA":"خرائط Bing","az-Latn-AZ":"Bing xəritələr","bg-BG":"Карти на Bing","bs-Latn-BA":"Bing karte","ca-ES":"Mapes del Bing","cs-CZ":"Mapy Bing","cy-GB":"Mapiau Bing","da-DK":"Bing Kort","de-DE":"Bing Karten","el-GR":"Χάρτες Bing","en-US":"Bing maps","es-ES":"Mapas de Bing","et-EE":"Bingi kaardid","eu-ES":"Bing Mapak","fi-FI":"Bing-kartat","fr-FR":"Bing Cartes","ga-IE":"Mapaí Bing","gl-ES":"Mapas de Bing","he-IL":"מפות Bing","hi-IN":"Bing मानचित्र","hr-HR":"Bing karte","hu-HU":"Bing Térképek","id-ID":"Peta Bing","it-IT":"Bing Maps","ja-JP":"Bing 地図","kk-KZ":"Bing карталары","ko-KR":"Bing 지도","lt-LT":"„Bing“ žemėlapiai","lv-LV":"Bing kartes","mk-MK":"Мапи Bing","ms-MY":"Peta Bing","nb-NO":"Bing-kart","nl-NL":"Bing Kaarten","pl-PL":"Mapy Bing","prs-AF":"نقشه های Bing","pt-BR":"Bing Mapas","pt-PT":"Mapas Bing","qps-ploc":"[!!--##Ƀĭńġ màƿš##--!!]","qps-ploca":"Ɓįƞġ mȃƿš","ro-RO":"Hărți Bing","ru-RU":"Карты Bing","sk-SK":"Mapy Bing","sl-SI":"Zemljevidi Bing","sr-Cyrl-RS":"Bing мапе","sr-Latn-RS":"Bing mape","sv-SE":"Bing-kartor","th-TH":"Bing Maps","tr-TR":"Bing haritalar","uk-UA":"Карти Bing","vi-VN":"Bản đồ Bing","zh-CN":"必应地图","zh-TW":"Bing 地圖服務","en-GB":"Bing maps","en-NZ":"Bing maps","en-IE":"Bing maps","en-AU":"Bing maps"},"description":{"default":"Display a key location on a map","ar-SA":"عرض موقع رئيسي على خريطة","az-Latn-AZ":"Xəritədə əsas yeri göstərin","bg-BG":"Показване на ключово местоположение на картата","bs-Latn-BA":"Prikažite ključnu lokaciju na karti","ca-ES":"Mostreu una ubicació clau en un mapa","cs-CZ":"Umožňuje zobrazit důležitou polohu na mapě.","cy-GB":"Dangos lleoliad allweddol ar y map","da-DK":"Vis en nøgleplacering på et kort","de-DE":"Einen wichtigen Ort auf einer Karte anzeigen.","el-GR":"Εμφανίστε μια σημαντική τοποθεσία σε έναν χάρτη","en-US":"Display a key location on a map","es-ES":"Muestra la ubicación clave en un mapa.","et-EE":"Saate kaardil kuvada põhiasukoha","eu-ES":"Erakutsi kokaleku nagusi bat mapa batean","fi-FI":"Näytä merkittävä sijainti kartalla","fr-FR":"Affichez un emplacement clé sur une carte","ga-IE":"Taispeáin príomhshuíomh ar mhapa","gl-ES":"Fixa unha localización importante no mapa.","he-IL":"הצג מיקום מרכזי במפה","hi-IN":"मानचित्र पर कोई मुख्य स्थान प्रदर्शित करें","hr-HR":"Prikaz ključnog mjesta na karti","hu-HU":"Elsődleges hely megjelenítése a térképen","id-ID":"Tampilkan lokasi penting di peta","it-IT":"Visualizzare una posizione chiave su una mappa","ja-JP":"主要な場所をマップに表示します","kk-KZ":"Негізгі орынды картадан көрсету","ko-KR":"지도에서 핵심 위치를 표시합니다.","lt-LT":"Rodyti pagrindinę vietą žemėlapyje","lv-LV":"Parādīt svarīgu atrašanās vietu kartē","mk-MK":"Покажете клучна локација на мапата","ms-MY":"Papar lokasi utama pada peta","nb-NO":"Vis en viktig plassering på et kart","nl-NL":"Belangrijke locatie op een kaart weergeven","pl-PL":"Wyświetl kluczową lokalizację na mapie","prs-AF":"نمایش یک موقعیت کلیدی در نقشه","pt-BR":"Exiba um local importante em um mapa","pt-PT":"Apresentar uma localização importante num mapa","qps-ploc":"[!!--##Ɖıŝƿľāƴ ä ķĕɏ ļơćąţıőň ŏń ȃ māƥ##--!!]","qps-ploca":"Ɗīšƿłâŷ á ķȇɏ ŀőċáťĩōņ ŏň ā màƥ","ro-RO":"Afișați o locație cheie pe o hartă","ru-RU":"Отображение ключевого расположения на карте","sk-SK":"Umožňuje zobraziť na mape dôležitú polohu","sl-SI":"Prikaži ključno lokacijo na zemljevidu.","sr-Cyrl-RS":"Прикажите кључну локацију на мапи","sr-Latn-RS":"Prikažite ključnu lokaciju na mapi","sv-SE":"Visa en viktig plats på en karta","th-TH":"แสดงตำแหน่งที่ตั้งหลักบนแผนที่","tr-TR":"Önemli bir konumu haritada görüntüleyin","uk-UA":"Відображення ключового розташування на карті","vi-VN":"Hiển thị vị trí quan trọng trên bản đồ","zh-CN":"在地图上显示关键位置","zh-TW":"在地圖上顯示重要位置","en-GB":"Display a key location on a map","en-NZ":"Display a key location on a map","en-IE":"Display a key location on a map","en-AU":"Display a key location on a map"},"officeFabricIconFontName":"MapPin","iconImageUrl":null,"groupId":"cf066440-0614-43d6-98ae-0b31cf14c7c3","group":{"default":"Media and Content","ar-SA":"الوسائط والمحتويات","az-Latn-AZ":"Media və Məzmun","bg-BG":"Мултимедия и съдържание","bs-Latn-BA":"Mediji i sadržaj","ca-ES":"Fitxers multimèdia i contingut","cs-CZ":"Multimédia a obsah","cy-GB":"Cyfryngau a Chynnwys","da-DK":"Medier og indhold","de-DE":"Medien und Inhalt","el-GR":"Πολυμέσα και περιεχόμενο","en-US":"Media and Content","es-ES":"Contenido y elementos multimedia","et-EE":"Meediumid ja sisu","eu-ES":"Multimedia eta edukia","fi-FI":"Media ja sisältö","fr-FR":"Média et contenu","ga-IE":"Meáin agus inneachar","gl-ES":"Contido e elementos multimedia","he-IL":"מדיה ותוכן","hi-IN":"मीडिया और सामग्री","hr-HR":"Mediji i sadržaj","hu-HU":"Média és tartalom","id-ID":"Media dan Konten","it-IT":"Elementi multimediali e contenuto","ja-JP":"メディアおよびコンテンツ","kk-KZ":"Мультимедиа және контент","ko-KR":"미디어 및 콘텐츠","lt-LT":"Medija ir turinys","lv-LV":"Multivide un saturs","mk-MK":"Медиуми и содржина","ms-MY":"Media dan Kandungan","nb-NO":"Medier og innhold","nl-NL":"Media en inhoud","pl-PL":"Multimedia i zawartość","prs-AF":"مطبوعات و محتوا","pt-BR":"Mídia e Conteúdo","pt-PT":"Multimédia e Conteúdo","qps-ploc":"[!!--##Mēđĩǻ ȃņđ Ĉōńťȇņť##--!!]","qps-ploca":"Měďīǻ ǻńď Ċōƞţȅƞţ","ro-RO":"Media și conținut","ru-RU":"Мультимедиа и контент","sk-SK":"Médiá a obsah","sl-SI":"Predstavnost in vsebina","sr-Cyrl-RS":"Медији и садржај","sr-Latn-RS":"Mediji i sadržaj","sv-SE":"Media och innehåll","th-TH":"สื่อและเนื้อหา","tr-TR":"Medya ve İçerik","uk-UA":"Мультимедіа та вміст","vi-VN":"Phương tiện và Nội dung","zh-CN":"媒体和内容","zh-TW":"媒體及內容","en-GB":"Media and Content","en-NZ":"Media and Content","en-IE":"Media and Content","en-AU":"Media and Content"},"properties":{"pushPins":[],"maxNumberOfPushPins":1,"shouldShowPushPinTitle":true,"zoomLevel":12,"mapType":"road"}}],"disabledOnClassicSharepoint":false,"searchablePropertyNames":null,"linkPropertyNames":null,"imageLinkPropertyNames":null,"hiddenFromToolbox":false,"supportsFullBleed":false,"requiredCapabilities":{"BingMapsKey":true},"isolationLevel":"None","version":"1.2.0","alias":"BingMapWebPart","preloadComponents":null,"isInternal":true,"loaderConfig":{"internalModuleBaseUrls":["https://spoprod-a.akamaihd.net/files/"],"entryModuleId":"sp-bing-map-webpart-bundle","exportName":null,"scriptResources":{"sp-bing-map-webpart-bundle":{"type":"localizedPath","shouldNotPreload":false,"id":null,"version":null,"failoverPath":null,"path":null,"defaultPath":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_default_f9778f1c0ab9952745886d1d605776ef.js","paths":{"ar-SA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","ar":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","tzm-Latn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","ku":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","syr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ar-sa_17aae2fc1378bb9a605568e37393b143.js","az-Latn-AZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_az-latn-az_85fe3797228bbb31c90827117a1e88b6.js","bg-BG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_bg-bg_c18a087f3cd05dfb53d2a764c84ee8e5.js","bs-Latn-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_bs-latn-ba_6d0592293c02851de2b8d6426de41275.js","ca-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ca-es_67b4773f20c230a76857bd60bf748dac.js","cs-CZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cs-cz_64b585926105d1854c19a591a96d4742.js","cy-GB":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cy-gb_d65e5b362d048891264360153e8b3925.js","cy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_cy-gb_d65e5b362d048891264360153e8b3925.js","da-DK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","fo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","kl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_da-dk_8891141e504d7dc9d7a37df5556d0c6a.js","de-DE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","de":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","dsb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","rm":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","hsb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_de-de_f25c96f5c7c5e70f9737413f493740b6.js","el-GR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_el-gr_fb1d150ca6185cc1aeb3077b88b19884.js","el":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_el-gr_fb1d150ca6185cc1aeb3077b88b19884.js","en-US":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","bn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","chr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","dv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","div":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","en":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","fil":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","haw":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","iu":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","lo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","moh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-us_f9778f1c0ab9952745886d1d605776ef.js","es-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","gn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","quz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","es":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","ca-ES-valencia":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_es-es_672cfd4f7e90302dd5bcf5bc7e569445.js","et-EE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_et-ee_4843ca005a2b35f4d37641aa010c9841.js","eu-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_eu-es_9a9719aea4f15cae349923e677971eb9.js","fi-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","sms":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","se-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","se-Latn-FI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fi-fi_9e81af24aa4f0ea2b7aa7cb45a0cdf47.js","fr-FR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","gsw":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","br":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","tzm-Tfng":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","co":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","fr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","ff":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","lb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","mg":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","oc":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","zgh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","wo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_fr-fr_36289327bc808b24faf87af94c4282e4.js","ga-IE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ga-ie_293592b7ef4495b6f2ac24ee7209a4ba.js","gl-ES":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_gl-es_73fd41fe859fc930e937f011dd817df0.js","he-IL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_he-il_99de68842371a4fc44f2c0b6ff9a5728.js","hi-IN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hi-in_4b5ca5e56fe8161732825c681204a7b8.js","hi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hi-in_4b5ca5e56fe8161732825c681204a7b8.js","hr-HR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hr-hr_f390168a1e0ecbe1449b3671bb8d40e4.js","hu-HU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_hu-hu_45735154b266d9a5c4c22d5f58ff357b.js","id-ID":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_id-id_16c289d93cc0beeff4bc67251a3eee1f.js","jv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_id-id_16c289d93cc0beeff4bc67251a3eee1f.js","it-IT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_it-it_01d4cf9abd13c78a05176b5603fcd910.js","it":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_it-it_01d4cf9abd13c78a05176b5603fcd910.js","ja-JP":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ja-jp_1e2c2f23c172b5b6cb58db6a93c22377.js","kk-KZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_kk-kz_b964c6292da4da296886bdb4333612af.js","ko-KR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ko-kr_732fca9752fe7e578a07302ec1664b60.js","lt-LT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_lt-lt_bf02cafe5c3eb67d6c226b57dec12e97.js","lv-LV":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_lv-lv_b1663a1f172b8fc522347238c799f7fd.js","mk-MK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_mk-mk_e98c9da753231452008db194cb735a2d.js","ms-MY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ms-my_2f6050a7c3ce3b6ffbb47689a4abaf29.js","ms":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ms-my_2f6050a7c3ce3b6ffbb47689a4abaf29.js","nb-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","no":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nb":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","smj-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","smj-Latn-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","se-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","se-Latn-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","sma-Latn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","sma-NO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nb-no_ed6f901c7b0f38d5e32dd4f3527a8ee4.js","nl-NL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","nl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","fy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_nl-nl_5e890532e3ff623812d8b0342e72ce9c.js","pl-PL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pl-pl_605a6ac656d59c71bfd6db4678c92ff2.js","prs-AF":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","gbz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","ps":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_prs-af_86e7fa58790e6cb3d3d7682793cc0d87.js","pt-BR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-br_9008a60143e5a19971076daced802a6e.js","pt-PT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-pt_224b200ba7df98bc7e2dfdad6762b7d9.js","pt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_pt-pt_224b200ba7df98bc7e2dfdad6762b7d9.js","qps-ploc":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_qps-ploc_7cc2e80271e05822b85637079fbbc8e6.js","qps-ploca":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_qps-ploca_cc832fbdf67d10b7072faeaecf7a48f7.js","ro-RO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ro-ro_4d43cca645941d769735433b76b88c92.js","ro":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ro-ro_4d43cca645941d769735433b76b88c92.js","ru-RU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ru":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ba":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","be":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","ky":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","mn":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","sah":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tg":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","tk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_ru-ru_90768dad2a3d28bb6be4e85842e8fe69.js","sk-SK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sk-sk_43b75662b363ee2797c81026449182af.js","sk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sk-sk_43b75662b363ee2797c81026449182af.js","sl-SI":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sl-si_4fad5739556858cca3f39917cb8b0cbd.js","sl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sl-si_4fad5739556858cca3f39917cb8b0cbd.js","sr-Cyrl-RS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-cyrl-rs_c2b5a41feb0599111c17612d389880cb.js","sr-Cyrl":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-cyrl-rs_c2b5a41feb0599111c17612d389880cb.js","sr-Latn-RS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-latn-rs_8ccc6f642938bf9e96f231e693a28fd2.js","sr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sr-latn-rs_8ccc6f642938bf9e96f231e693a28fd2.js","sv-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","smj":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","se":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sv":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sma-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","sma-Latn-SE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_sv-se_dd07ff103d28b74d614c836cd86dd450.js","th-TH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_th-th_04b7da4b93097cb2e7f295e1afd80d19.js","th":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_th-th_04b7da4b93097cb2e7f295e1afd80d19.js","tr-TR":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_tr-tr_d81724d2147360856bba8c7de5acc476.js","tr":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_tr-tr_d81724d2147360856bba8c7de5acc476.js","uk-UA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_uk-ua_9365a5370d64d5650444397ab89e4b5c.js","vi-VN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_vi-vn_e45cb5b2cb1a7bcb517fb17e04188820.js","vi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_vi-vn_e45cb5b2cb1a7bcb517fb17e04188820.js","zh-CN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","zh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","mn-Mong":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","bo":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","ug":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","ii":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-cn_5ec450459292137d1a9205b0e54730fc.js","zh-TW":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-HK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-CHT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-Hant":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","zh-MO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_zh-tw_235c724e9df87fbacc35d582988c0b55.js","en-GB":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sq":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","am":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","hy":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","mk":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","bs":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","my":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","dz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-CY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-EG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-IL":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-IS":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-JO":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-KE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-KW":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-MK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-MT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-PK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-QA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-SA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-LK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-AE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-VN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","is":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","km":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","kh":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","mt":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","fa":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","gd":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sr-Cyrl-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sr-Latn-BA":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","sd":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","si":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","so":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","ti-ET":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","uz":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-gb_f9778f1c0ab9952745886d1d605776ef.js","en-NZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-nz_f9778f1c0ab9952745886d1d605776ef.js","en-IE":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-ie_f9778f1c0ab9952745886d1d605776ef.js","en-AU":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-SG":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-HK":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-MY":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-PH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-TT":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-AZ":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-BH":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-BN":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","en-ID":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js","mi":"sp-client-prod_2018-06-15.009/sp-bing-map-webpart-bundle_en-au_f9778f1c0ab9952745886d1d605776ef.js"},"globalName":null,"globalDependencies":null},"react":{"type":"component","shouldNotPreload":false,"id":"0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d","version":"15.6.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/office-ui-fabric-react-bundle":{"type":"component","shouldNotPreload":false,"id":"02a01e42-69ab-403d-8a16-acd128661f8e","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/load-themed-styles":{"type":"component","shouldNotPreload":false,"id":"229b8d08-79f3-438b-8c21-4613fc877abd","version":"0.1.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/odsp-utilities-bundle":{"type":"component","shouldNotPreload":false,"id":"cc2cc925-b5be-41bb-880a-f0f8030c6aff","version":"4.1.3","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-telemetry":{"type":"component","shouldNotPreload":false,"id":"8217e442-8ed3-41fd-957d-b112e841286a","version":"0.2.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-core-library":{"type":"component","shouldNotPreload":false,"id":"7263c7d0-1d6a-45ec-8d85-d4d1d234171b","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-webpart-base":{"type":"component","shouldNotPreload":false,"id":"974a7777-0990-4136-8fa6-95d80114c2e0","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-webpart-shared":{"type":"component","shouldNotPreload":false,"id":"914330ee-2df2-4f6e-a858-30c23a812408","version":"0.1.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-component-utilities":{"type":"component","shouldNotPreload":false,"id":"8494e7d7-6b99-47b2-a741-59873e42f16f","version":"0.2.1","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"react-dom":{"type":"component","shouldNotPreload":false,"id":"aa0a46ec-1505-43cd-a44a-93f3a5aa460a","version":"15.6.2","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-lodash-subset":{"type":"component","shouldNotPreload":false,"id":"73e1dc6c-8441-42cc-ad47-4bd3659f8a3a","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@microsoft/sp-diagnostics":{"type":"component","shouldNotPreload":false,"id":"78359e4b-07c2-43c6-8d0b-d060b4d577e8","version":"1.6.0","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null},"@ms/sp-bingmap":{"type":"component","shouldNotPreload":false,"id":"ab22169a-a644-4b69-99a2-4295eb0f633c","version":"0.0.1","failoverPath":null,"path":null,"defaultPath":null,"paths":null,"globalName":null,"globalDependencies":null}}},"manifestVersion":2,"id":"e377ea37-9047-43b9-8cdb-a761be2f8e09","componentType":"WebPart"}',
		ManifestType: 1,
		Name: 'e377ea37-9047-43b9-8cdb-a761be2f8e09',
		Status: 0
	};

	// The Mock canvas has 3 section (1st has 1 column, 2nd has 2 columns, 3rd has 3 columns)
	// column 1 of section 2 have 3 Text controls to test the order argument
	const clientSidePageCanvasContent = `"<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;12,&quot;zoneIndex&quot;&#58;1,&quot;layoutIndex&quot;&#58;1&#125;&#125;"></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;b39b599e-4b8b-4f97-b262-a5e08fab9253&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;2,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;6,&quot;layoutIndex&quot;&#58;1&#125;,&quot;editorType&quot;&#58;&quot;CKEditor&quot;&#125;"><div data-sp-rte=""><p>Text 1</p>
  </div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;e8ddd136-90e9-47e7-97e3-64bda61f72f1&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;2,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;2,&quot;layoutIndex&quot;&#58;1&#125;,&quot;editorType&quot;&#58;&quot;CKEditor&quot;&#125;"><div data-sp-rte=""><p>Text 2</p>
  </div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,&quot;id&quot;&#58;&quot;aa393bd8-43df-414f-ad0d-9eb8eac27334&quot;,&quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;2,&quot;sectionIndex&quot;&#58;1,&quot;controlIndex&quot;&#58;3,&quot;layoutIndex&quot;&#58;1&#125;,&quot;editorType&quot;&#58;&quot;CKEditor&quot;&#125;"><div data-sp-rte=""><p>Text 3</p>
  </div></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;6,&quot;zoneIndex&quot;&#58;2,&quot;layoutIndex&quot;&#58;1&#125;&#125;"></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;1,&quot;sectionFactor&quot;&#58;4,&quot;zoneIndex&quot;&#58;3,&quot;layoutIndex&quot;&#58;1&#125;&#125;"></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;2,&quot;sectionFactor&quot;&#58;4,&quot;zoneIndex&quot;&#58;3,&quot;layoutIndex&quot;&#58;1&#125;&#125;"></div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" data-sp-controldata="&#123;&quot;displayMode&quot;&#58;2,&quot;position&quot;&#58;&#123;&quot;sectionIndex&quot;&#58;3,&quot;sectionFactor&quot;&#58;4,&quot;zoneIndex&quot;&#58;3,&quot;layoutIndex&quot;&#58;1&#125;&#125;"></div></div>"`;
  const emptyPage = '<div></div>';
	beforeEach(() => {
		log = [];
		commandContext = {
			requestContext: {
				accessToken: '',
				requestDigest: ''
			},
			debug: false,
			verbose: false,
			webUrl: 'https://contoso.sharepoint.com/sites/team-a',
			pageName: 'page.aspx',
			log: (message: any) => {
				log.push(message);
			}
    };
    
    loggerSpy = sinon.spy(commandContext, 'log');

		sinon.stub(request, 'get').callsFake((opts) => {
			// Fake the available Client Side WebParts (Limited to the BingMap WebPart)
			if (opts.url.indexOf(`/_api/web/getclientsidewebparts()`) > -1) {
				const fakeResponse = {
					value: [ bingMapWebPartDefinition ]
				};
				return Promise.resolve(fakeResponse);
			}

			// Fake the Get client side page
			if (
				opts.url.indexOf(
					`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')?$expand=ListItemAllFields/ClientSideApplicationId`
				) > -1
			) {
				return Promise.resolve({
					ListItemAllFields: {
						CommentsDisabled: false,
						FileSystemObjectType: 0,
						Id: 1,
						ServerRedirectedEmbedUri: null,
						ServerRedirectedEmbedUrl: '',
						ContentTypeId: '0x0101009D1CB255DA76424F860D91F20E6C41180062FDF2882AB3F745ACB63105A3C623C9',
						FileLeafRef: 'Home.aspx',
						ComplianceAssetId: null,
						WikiField: null,
						Title: 'Home',
						ClientSideApplicationId: 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec',
						PageLayoutType: 'Home',
						CanvasContent1: clientSidePageCanvasContent,
						BannerImageUrl: {
							Description: '/_layouts/15/images/sitepagethumbnail.png',
							Url: 'https://contoso.sharepoint.com/_layouts/15/images/sitepagethumbnail.png'
						},
						Description: 'Lorem ipsum Dolor samet Lorem ipsum',
						PromotedState: null,
						FirstPublishedDate: null,
						LayoutWebpartsContent: null,
						AuthorsId: null,
						AuthorsStringId: null,
						OriginalSourceUrl: null,
						ID: 1,
						Created: '2018-01-20T09:54:41',
						AuthorId: 1073741823,
						Modified: '2018-04-12T12:42:47',
						EditorId: 12,
						OData__CopySource: null,
						CheckoutUserId: null,
						OData__UIVersionString: '7.0',
						GUID: 'edaab907-e729-48dd-9e73-26487c0cf592'
					},
					CheckInComment: '',
					CheckOutType: 2,
					ContentTag: '{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25,1',
					CustomizedPageStatus: 1,
					ETag: '"{E82A21D1-CA2C-4854-98F2-012AC0E7FA09},25"',
					Exists: true,
					IrmEnabled: false,
					Length: '805',
					Level: 1,
					LinkingUri: null,
					LinkingUrl: '',
					MajorVersion: 7,
					MinorVersion: 0,
					Name: 'home.aspx',
					ServerRelativeUrl: '/sites/team-a/SitePages/home.aspx',
					TimeCreated: '2018-01-20T08:54:41Z',
					TimeLastModified: '2018-04-12T10:42:46Z',
					Title: 'Home',
					UIVersion: 3584,
					UIVersionLabel: '7.0',
					UniqueId: 'e82a21d1-ca2c-4854-98f2-012ac0e7fa09'
				});
			}

			return Promise.reject('Invalid request');
		});

		sinon.stub(request, 'post').callsFake((opts) => {
			// Fake the POST to update ClientSidePage
			if (
				opts.url.indexOf(
					`/_api/web/getfilebyserverrelativeurl('/sites/team-a/sitepages/page.aspx')/ListItemAllFields`
				) > -1
			) {
				return Promise.resolve({});
			}

			return Promise.reject('Invalid request');
		});
	});

	afterEach(() => {
		Utils.restore([ request.post, request.get, StandardWebPartUtils.getWebPartId ]);
	});

	it('gets a Client Side page object instance', (done) => {
		commandContext.pageName = 'page.aspx';
		ClientSidePageCommandHelper.getClientSidePage(commandContext).then((clientSidePage) => {
			// Make sure the Client Side Page object has the required methods
			assert(clientSidePage.addSection && clientSidePage.toHtml);
			done();
		});
	});

	it('gets a Client Side WebPart object instance', (done) => {
		ClientSidePageCommandHelper.getWebPartInstance(commandContext, bingMapWebPartId)
			.then((webPart) => {
				const actualWebPartId = webPart.getControlData().webPartId;
				assert(actualWebPartId, bingMapWebPartId);
				done();
			})
			.catch((e) => {
				done(e);
			});
	});

	it('gets a standard Client Side WebPart object instance from its name ', (done) => {
		ClientSidePageCommandHelper.getStandardWebPartInstance(commandContext, bingMapWebPartName)
			.then((webPart) => {
				const expectedJsonDataEnd =
					'&quot;webPartId&quot;&#58;&quot;e377ea37-9047-43b9-8cdb-a761be2f8e09&quot;&#125;';
				assert(webPart.jsonData.endsWith(expectedJsonDataEnd));
				done();
			})
			.catch((e) => {
				done(e);
			});
	});

	it('fails when a standard Client Side WebPart definition could not be found from API ', (done) => {
		const fakeWebPartId = 'XXXX';
		sinon.stub(StandardWebPartUtils, 'getWebPartId').callsFake((opts) => fakeWebPartId);

		ClientSidePageCommandHelper.getStandardWebPartInstance(commandContext, bingMapWebPartName)
			.then(() => {
				assert.fail('WebPart should not be returned in this case');
				done();
			})
			.catch((error) => {
				try {
					assert.equal(
						JSON.stringify(error),
						JSON.stringify(new Error(`There is no available WebPart with Id '${fakeWebPartId}'`))
					);
					done();
				} catch (e) {
					done(e);
				}
			});
	});

	it('adds a Client Side WebPart to a modern page', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		const webPart = ClientSideWebpart.fromComponentDef(bingMapWebPartDefinition);
		ClientSidePageCommandHelper.addWebPartToPage(page, webPart, commandContext);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[0].columns[0].getControl(0);
		const controlData = control.controlData as ClientSideControlData;
		assert.equal(controlData.webPartId, bingMapWebPartId);
		done();
	});

	it('adds a Client Side WebPart to a modern page at section 2 column 2', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		const webPart = ClientSideWebpart.fromComponentDef(bingMapWebPartDefinition);
		ClientSidePageCommandHelper.addWebPartToPage(page, webPart, commandContext, 2, 2);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[1].columns[1].getControl(0);
    const controlData = control.controlData as ClientSideControlData;
    assert.equal(controlData.webPartId, bingMapWebPartId);
		done();
  });
  
  it('adds a Client Side WebPart to a modern page at a specific order', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		const webPart = ClientSideWebpart.fromComponentDef(bingMapWebPartDefinition);
		ClientSidePageCommandHelper.addWebPartToPage(page, webPart, commandContext, 2, 1, 2);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[1].columns[0].getControl(1);
    const controlData = control.controlData as ClientSideControlData;
    assert.equal(controlData.webPartId, bingMapWebPartId);
		done();
	});

	it('adds a Client Side WebPart to a modern page at section 1 column 1 with properties', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		const webPart = ClientSideWebpart.fromComponentDef(bingMapWebPartDefinition);
		ClientSidePageCommandHelper.addWebPartToPage(page, webPart, commandContext, 1, 1, 0, '{"Title":"Location"}');

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[0].columns[0].getControl(0);
		const controlData = control.controlData as ClientSideControlData;
    assert.equal(controlData.webPartId, bingMapWebPartId);
		done();
  });

  it('adds a Client Side Text to a modern page', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		ClientSidePageCommandHelper.addTextToPage(page, 'Foobar', commandContext);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[0].columns[0].getControl(0) as ClientSideText;
    assert.equal(control.text, '<p>Foobar</p>');
		done();
  });

  it('handles empty page', (done) => {
    const page = ClientSidePage.fromHtml(emptyPage);
    ClientSidePageCommandHelper.addTextToPage(page, 'Foobar', commandContext);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[0].columns[0].getControl(0) as ClientSideText;
    assert.equal(control.text, '<p>Foobar</p>');
		done();
  });

  it('adds a Client Side Text to a modern page with specific section and column', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		ClientSidePageCommandHelper.addTextToPage(page, 'Foobar', commandContext, 2, 2);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[1].columns[1].getControl(0) as ClientSideText;
    assert.equal(control.text, '<p>Foobar</p>');
		done();
  });

  it('adds a Client Side Text to a modern page at a specific order', (done) => {
		const page = ClientSidePage.fromHtml(clientSidePageCanvasContent);
		ClientSidePageCommandHelper.addTextToPage(page, 'Foobar', commandContext, 2, 1, 2);

		const updatedPage = ClientSidePage.fromHtml(page.toHtml());
		const control = updatedPage.sections[1].columns[0].getControl(1) as ClientSideText;
    assert.equal(control.text, '<p>Foobar</p>');
		done();
  });
  
  it ('handles debug mode', (done) => {
    commandContext.debug = true;
    ClientSidePageCommandHelper.getClientSidePage(commandContext).then((clientSidePage) => {
			// Make sure the say() method is called in verbose mode with text 'Executing web request...'
			assert(loggerSpy.calledWith('Executing web request...'));
			done();
		});
  });

  it ('handles verbose mode', (done) => {
    commandContext.verbose = true;
    ClientSidePageCommandHelper.getClientSidePage(commandContext).then((clientSidePage) => {
			// Make sure the say() method is called in verbose mode with text 'Retrieving information about the page...'
			assert(loggerSpy.calledWith('Retrieving information about the page...'));
			done();
		});
  });
});
