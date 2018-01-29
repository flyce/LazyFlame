@echo off
set /p input=要处理的周数:
set /a week=%input%-1
echo 从 例会准备材料\%input%\ 复制 维修工程部第%week%周工作周报.xlsx 到 LazyFlame\public
copy 例会准备材料\%input%\维修工程部第%week%周工作周报.xlsx LazyFlame\public\doc.xlsx
cd LazyFlame
echo 打开浏览器，访问localhost:3000
nodemon ./bin/www