# docOnline
edit docx template online html  
now support content input and table input  
template input must be empty underlined, table first row must colum names



# Usage
- `pip install pydocx python-docx flask` 
- `python app.py` or other method develop flask app

# Router
- post /upload   
``` 
request:
curl --request POST \
  --url http://127.0.0.1:5000/upload \
  --header 'content-type: multipart/form-data' \
  --form 'file=@file.docx'
response:
{
	"datas": [
		"item3"
        ],
	"id": "250c36f5c081450a3c2dc9f0f2d371f5"
}
```
- get /html/@id  
```
<html>
....
</html>
```
- post /docx/@id  
```
request:  
curl --request POST \
  --url http://127.0.0.1:5000/docx/250c36f5c081450a3c2dc9f0f2d371f5 \
  --header 'content-type: application/json' \
  --data '{
    "tables": [],
    "datas":[
        "item1"
    ]
}'
response:  
filename=250c36f5c081450a3c2dc9f0f2d371f5.docx;attachment;
```
- get /docx/@id 

```
alarm: the id must posted 
filename=250c36f5c081450a3c2dc9f0f2d371f5.docx;attachment;
```
