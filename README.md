# excel op by phpexcel

### 安裝

先在composer.json中增加：

```json
"repositories": [
    ...,
    {
        "type": "vcs",
        "url": "https://github.com/exinfinite/Excel.git"
    }
]
```
```php
composer require Exinfinite\Excel
```

### 使用

```php
use Exinfinite\Excel;

$excel = new Excel();
```