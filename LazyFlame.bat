@echo off
set /p input=Ҫ���������:
set /a week=%input%-1
echo �� ����׼������\%input%\ ���� ά�޹��̲���%week%�ܹ����ܱ�.xlsx �� LazyFlame\public
copy ����׼������\%input%\ά�޹��̲���%week%�ܹ����ܱ�.xlsx LazyFlame\public\doc.xlsx
cd LazyFlame
echo �������������localhost:3000
nodemon ./bin/www