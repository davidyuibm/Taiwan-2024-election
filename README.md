

這是針對台灣 2024 年總統選舉, 選民在各投票所三大黨支持率排行榜, IDE 環境為 pycharm.

因為每次只針對一黨, 預設是對民進黨, df2 = df.assign(green=(df['Unnamed: 4']/ df['Unnamed: 6']))

對國民黨, 可以改成 df2 = df.assign(green=(df['Unnamed: 5']/ df['Unnamed: 6']))

對民眾黨, 可以改成 df2 = df.assign(green=(df['Unnamed: 3']/ df['Unnamed: 6']))

所有各縣市得票結果 *.xlsx, 是從台灣選舉委會下載. https://db.cec.gov.tw/ElecTable/Election/ElecTickets?dataType=tickets&typeId=ELC&subjectId=P0&legisId=00&themeId=4d83db17c1707e3defae5dc4d4e9c800&dataLevel=N&prvCode=00&cityCode=000&areaCode=00&deptCode=000&liCode=0000
