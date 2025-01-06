import pychrome

browser = pychrome.Browser(url="http://127.0.0.1:9222")
targets = browser.call_method("Target.getTargets")

for target in targets['targetInfos']:
    print(f"Target ID: {target['targetId']}, Type: {target['type']}, Title: {target['title']}")
