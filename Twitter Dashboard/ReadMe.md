
The Github READme.md Twitter Dashboard
The Meltwater Platform enables the consumption of many data sources including Social media. Examples include Twitter, Facebook, Instagram, Pinterst. News media include the New York Times a recent addition for Meltwater customers.
Meltwater states on its website, as at 18th December 2021: 
“We offer comprehensive media monitoring and analysis across online news, social media, print, broadcast, and podcasts, capturing more content and conversations than anyone else in the industry. Our suite of products aims to solve the problems faced by modern PR, communications, and marketing professionals.”

<img width="320" alt="Meltwater Webpage Intro" src="https://user-images.githubusercontent.com/47678539/146658331-886b3c13-1185-496e-9096-40126dfd412c.png">

 
The accompanying YouTube video (Video title: “Building a Twitter analytics dashboard using the Meltwater platform”), to this Github repo, is split into four parts as follows:
1.	Excel Dashboard intro.
2.	Meltwater Platform Intro and REST API generation.
3.	Power Automate Flow walkthrough “Meltwater Twitter Covid to Excel”.
4.	Excel Script walkthrough “InputCSVMeltwater”.

This demo shows how to integrate the Meltwater platform with the Office 365 platform utilising Power Automate, Excel Online and Excel Office Scripts. The dataset used for the demo was the count of Twitter comments containing the word “covid” listed by country of origin over a tumbling window of 5 minutes. 
The Excel workbook and Office Script are included within this repo. 

<img width="960" alt="Excel Dashboard Screenshot" src="https://user-images.githubusercontent.com/47678539/146658408-5d03aaa2-06c3-4526-b3ab-80841332a58e.PNG">


An example of the Meltwater dynamically Generated REST API URL
curl -X GET \ 
	--url "https://api.meltwater.com/v3/analytics/17302062/top_locations?start=2021-12-16T00:00:00&end=2021-12-17T00:00:00&tz=Europe/London&source=twitter&size=100&level=country" \ 
	--header "Accept: application/json" \ 
	--header "apikey: **********"

The concat function used in the Power Automate Flow HTTP action as demoed in the YouTube video:
concat('https://api.meltwater.com/v3/analytics/11964917/top_locations?start=',outputs('Compose-StartDate'),'T',outputs('Compose-StartTime'), '&end=',outputs('Compose-EndDate'),'T',outputs('Compose-EndTime'),'&tz=Europe/London&source=news&size=100&level=country')


The response data from Meltwater:
Count of Twitter comments containing the word “covid” – table format
location_name	location_id	document_count	Percentage
United States	us	67	30.88
Canada	ca	10	4.61
United Kingdom	gb	6	2.76
Australia	au	2	0.92
Brazil	br	2	0.92
China	cn	2	0.92
France	fr	2	0.92
Ireland	ie	2	0.92
Nigeria	ng	2	0.92
Peru	pe	2	0.92
Russia	ru	2	0.92
Bolivia	bo	1	0.46
Switzerland	ch	1	0.46
Colombia	co	1	0.46
Georgia	ge	1	0.46
Ghana	gh	1	0.46
Indonesia	id	1	0.46
Iran	ir	1	0.46
Italy	it	1	0.46
South Korea	kr	1	0.46
Sri Lanka	lk	1	0.46
Mexico	mx	1	0.46
Netherlands	nl	1	0.46
New Zealand	nz	1	0.46
Panama	pa	1	0.46
Sweden	se	1	0.46
Singapore	sg	1	0.46
Thailand	th	1	0.46
Tunisia	tn	1	0.46


The Power Automate Flow
 
![TwitterMeltwaterFlow](https://user-images.githubusercontent.com/47678539/146658461-9b503fc9-055c-40e3-8648-5da1d91bd53f.png)


**The Excel Script used in the Flow:**

...
function main(workbook: ExcelScript.Workbook, csvinput:string) {
  //console.log('defns : ' + defns)
 // let csvinput1:string = "location_name, location_id, document_count, percentage\r\nPoland, pl, 2020, 29.27\r\nGermany, de, 1279, 18.53\r\nGreece, gr, 526, 7.62\r\nSpain, es, 480, 6.95\r\nUnited Kingdom, gb, 381, 5.52\r\nItaly, it, 292, 4.23\r\nRomania, ro, 214, 3.1\r\nTurkey, tr, 145, 2.1\r\nSouth Africa, za, 125, 1.81\r\nSwitzerland, ch, 119, 1.72\r\nRussia, ru, 106, 1.54\r\nFrance, fr, 102, 1.48\r\nNetherlands, nl, 92, 1.33\r\nCroatia, hr, 88, 1.27\r\nNigeria, ng, 78, 1.13\r\nPortugal, pt, 69, 1\r\nAustria, at, 67, 0.97\r\nBelgium, be, 63, 0.91\r\nBulgaria, bg, 52, 0.75\r\nLatvia, lv, 47, 0.68\r\nGhana, gh, 44, 0.64\r\nUkraine, ua, 43, 0.62\r\nIsrael, il, 36, 0.52\r\nDenmark, dk, 35, 0.51\r\nCyprus, cy, 33, 0.48\r\nCzechia, cz, 28, 0.41\r\nIreland, ie, 27, 0.39\r\nSlovakia, sk, 27, 0.39\r\nHungary, hu, 25, 0.36\r\nLuxembourg, lu, 22, 0.32\r\nMoldova, md, 22, 0.32\r\nFinland, fi, 20, 0.29\r\nPakistan, pk, 19, 0.28\r\nAlbania, al, 16, 0.23\r\nLithuania, lt, 13, 0.19\r\nSweden, se, 12, 0.17\r\nUganda, ug, 12, 0.17\r\nEgypt, eg, 9, 0.13\r\nSlovenia, si, 9, 0.13\r\nBosnia & Herzegovina, ba, 8, 0.12\r\nJordan, jo, 8, 0.12\r\nKazakhstan, kz, 8, 0.12\r\nNorway, no, 8, 0.12\r\nMorocco, ma, 7, 0.1\r\nSaudi Arabia, sa, 6, 0.09\r\nUruguay, uy, 6, 0.09\r\nAzerbaijan, az, 4, 0.06\r\nIceland, is, 4, 0.06\r\nAngola, ao, 3, 0.04\r\nGeorgia, ge, 3, 0.04\r\nIran, ir, 3, 0.04\r\nCambodia, kh, 3, 0.04\r\nTunisia, tn, 3, 0.04\r\nTanzania, tz, 3, 0.04\r\nZambia, zm, 3, 0.04\r\nZimbabwe, zw, 3, 0.04\r\nBahrain, bh, 2, 0.03\r\nCentral African Republic, cf, 2, 0.03\r\nEstonia, ee, 2, 0.03\r\nMontenegro, me, 2, 0.03\r\nMacedonia, mk, 2, 0.03\r\nNamibia, na, 2, 0.03\r\nFrench Polynesia, pf, 2, 0.03\r\nRwanda, rw, 2, 0.03\r\nArmenia, am, 1, 0.01\r\nBotswana, bw, 1, 0.01\r\nCongo - Brazzaville, cg, 1, 0.01\r\nCameroon, cm, 1, 0.01\r\nAlgeria, dz, 1, 0.01\r\nGambia, gm, 1, 0.01\r\n"
  let Sheet1 = workbook.getWorksheet('Sheet1')
  let string1 = csvinput.toString()
  if (string1.length == 0) {
    return
  }
//Sheet1.getCell(0, 0).setValue(csvinput)
//Headers remain
//See Script: ImportingArrayRowByRow
//Clear existing data
Sheet1.getRange('A2:D40').clear(ExcelScript.ClearApplyTo.contents)
Sheet1.getRange('A41:D80').clear(ExcelScript.ClearApplyTo.contents)

let inputArray = csvinput.split('\r\n')
  //console.log('inputArray:' + inputArray)
let leninputArray:number =  inputArray.length
  //console.log('leninputArray:'+leninputArray)
  for (let k = 0; k < leninputArray+1; k++) {
    Sheet1.getCell(k, 0).setValue(inputArray[k])
    //Split column A by comma
    let inputArrayKStr = Sheet1.getCell(k, 0).getValue().toString() 
    //let inputArrayKStr = inputArray[k].toString()
    let colAArray = inputArrayKStr.split(',')
    //console.log('colAArray:' +colAArray)
    let lenAArray = colAArray.length 
    for (let L = 0; L < lenAArray + 1; L++) {
      Sheet1.getCell(k, L).setValue(colAArray[L])
    //console.log(defnsArray[k])
  }
  }
  Sheet1.getRange("A:K").getFormat().autofitColumns()
  //Last Refresh:
  let date: Date = new Date();
  Sheet1.getCell(25, 5).setValue('Last Refresh: ' + date)
  console.log('Last Refresh: ' + date)

  return
  }
...


