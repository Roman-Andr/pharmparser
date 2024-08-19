# Get started
```
pip install -r requirements.txt
```
## Make test request
You can use program like Postman
### Example cookies:
```json
"PHPSESSID": "...",
"_csrf": "...",
"_ga": "...",
"_ga_S6LL4MRH46": "...",
"regionId": "...",
"lim-result": "some number"
```
### Example data
```json
"_csrf": "...",
"id": "target id",
"page": "0",
"sort": "name",
"sort_type": "asc"
```
Then insert this to ParserEngine request call
