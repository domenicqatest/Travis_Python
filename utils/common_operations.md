Common Operations
=================

Hovering
--------

```
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains

driver = webdriver.Chome()
driver.get('http://example.com')

element_to_hover_over = driver.find_element_by_id("id_name")

hover = ActionChains(driver).move_to_element(element_to_hover_over)
hover.perform()
```

Working with tabs
-----------------

When a new tab is opened, and you want to close the tab and resume testing in the original tab:

```
driver.close()
driver.switch_to.window(driver.window_handles[0])
```