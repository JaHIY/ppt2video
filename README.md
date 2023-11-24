# PPT2Video

```
$ poetry config installer.modern-installation false
```

```
$ poetry install
```

```
$ poetry run python ppt2video convert -i .\input.pptx .\output.mp4
```

```
$ poetry run python ppt2video convert --dpi 300 --voice "zh-CN-YunxiNeural" -i .\input.pptx .\output.mp4
```

```
$ poetry run python ppt2video list-voices --language zh --gender female
```

```
$ poetry run python ppt2video list-voices --locale zh-CN --gender male
```
