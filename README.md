# Постраничный парсер данных

Парсер, реализованный через создание Firefox-драйверов в Selenium, каждый из которых при открытии собирает данные с новой страницы.
Сами драйверы легко заменяются на Chrome или другие популярные браузеры (см. документацию [Selenium](https://www.selenium.dev/documentation/en/webdriver/driver_requirements/))

Данные собираются через систему XPath, HTML-классы и построчного поиска. Собранные строки единообразных данных приводятся к финальному виду, после чего данные с каждой страницы построчно собираются в Excel таблицу. Страницы не содержащие обязательных данных - пропускаются.
