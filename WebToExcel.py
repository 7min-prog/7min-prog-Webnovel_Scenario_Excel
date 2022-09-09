from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import openpyxl as excel


def GetScenarioTextLines(url):
    # シナリオURLを開く
    driver.get(url)
    # シナリオ本文のinnerHTMLを<br>タグ区切り(行区切り)で配列に格納
    idName = driver.find_element_by_id("body")
    innerTextByLines = idName.get_attribute("innerHTML").split("<br>")
    # 整形済のTextを入れる配列を用意
    scenarioTextLines = []
    # 各行先頭の改行タグと空白を削除して配列に逐次追加
    for i, line in enumerate(innerTextByLines):
        scenarioTextLines.append(line.replace(" ", "").replace("\n", ""))
        print(scenarioTextLines[i])
    sleep(0.1)
    return scenarioTextLines


def SaveScenarioToExcel(section, urlDictionary):
    book = excel.Workbook()
    sheet = book.active
    sheet.title = section
    sheet["A1"] = "Command"
    sheet["B1"] = "Arg1"
    sheet["C1"] = "Arg2"
    sheet["D1"] = "Arg3"
    sheet["E1"] = "Arg4"
    sheet["F1"] = "Arg5"
    sheet["G1"] = "Arg6"
    sheet["H1"] = "Text"

    row = 2
    for url in urlDictionary.values():
        lines = GetScenarioTextLines(url)

        for line in lines:
            # コマンド行はA列に格納
            if any(map(line[:7].__contains__, ("BGM", "CG", "SE", "カットイン"))):
                sheet[f"A{row}"] = line
            # セリフ行は発言者をB列に格納し、セリフを「」付きでH列に格納
            elif "「" in line[:7]:
                speaker = line.split("「")[0]
                sheet[f"B{row}"] = speaker
                sheet[f"H{row}"] = line[len(speaker):]
            # 非発言テキストはそのままH列に格納
            else:
                sheet[f"H{row}"] = line
            row = row + 1

    book.save(f"ScenarioFiles/{section}.xlsx")


def GetUrlDictionarySec1():
    urlDictionary = {"1-S1": "https://notes.underxheaven.com/preview/fc87355c695b2e7c672b3f26f66bd46d",
                     "1-S2": "https://notes.underxheaven.com/preview/619b62b0ddbc273e62bbbf23071b821d",
                     "1-S3": "https://notes.underxheaven.com/preview/ee07b004f72dc40ed3b817800e5c7793",
                     "1-S4": "https://notes.underxheaven.com/preview/80df5fbb9efa3d86a91b8899550b390e",
                     "1-S5": "https://notes.underxheaven.com/preview/a41694540c860b5e902ec6881b31dfdb",
                     "1-S6": "https://notes.underxheaven.com/preview/4be39c380c00cf57275d4c12eaff54fe",
                     "1-S7": "https://notes.underxheaven.com/preview/1dbc6189cf17a8d426db2f50e621b193",
                     "1-S8": "https://notes.underxheaven.com/preview/63da6a44948b34fbe02f231904fffd96"
                     }
    return urlDictionary
    
def GetUrlDictionarySec2():
    urlDictionary = {"2-S1": "https://notes.underxheaven.com/preview/e44e00c5eb414325b926a9591f00bfa7",
                     "2-S2": "https://notes.underxheaven.com/preview/bb9d6fe0805beba89924ac04df5ef744",
                     "2-S3": "https://notes.underxheaven.com/preview/4e8fc734dac01c2b3cc59b02c0b8c0b3",
                     "2-S4": "https://notes.underxheaven.com/preview/09c8c4ff1982b6d3338f6a4e0a9b6be4",
                     "2-S5": "https://notes.underxheaven.com/preview/f3ccd83012a4b28e41c525dceb72f479",
                     "2-S6": "https://notes.underxheaven.com/preview/a803587bddd29bd751b86a720206b6b1",
                     "2-S7": "https://notes.underxheaven.com/preview/7b9a720bf0ecdec25f24b0886b2de900",
                     "2-S8": "https://notes.underxheaven.com/preview/e0f933ff04eb3ba04f69e9e90975d328",
                     "2-S9": "https://notes.underxheaven.com/preview/fe10566a38513da080fda711a0b55dff",
                     "2-S10": "https://notes.underxheaven.com/preview/591e6952970d92566f9bbe51e9ad4e57",
                     "2-S11": "https://notes.underxheaven.com/preview/ff78624972d14ce133a153722051cfed",
                     "2-S12": "https://notes.underxheaven.com/preview/44b6c6ef0462f0b85c0a225671bfdee3",
                     "2-S13": "https://notes.underxheaven.com/preview/fef33004df8b90cea70600981aa191cb",
                     "2-S14": "https://notes.underxheaven.com/preview/c996c68466c436fbdb045870210e091c",
                     "2-S15": "https://notes.underxheaven.com/preview/dc2de996f58812c212a00b8ad169ab7f",
                     "2-S16": "https://notes.underxheaven.com/preview/451bf1fd14b814185c172fc03fe13cb7"
                     }
    return urlDictionary


if __name__ == "__main__":
    # ChromeのバージョンにあったドライバをインストールしてからChromeを開く
    driver = webdriver.Chrome(ChromeDriverManager().install())
    SaveScenarioToExcel("Section1", urlDictionary = GetUrlDictionarySec1())
    SaveScenarioToExcel("Section2", urlDictionary = GetUrlDictionarySec2())

    driver.quit()


