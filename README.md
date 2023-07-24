
# Install

See [Yeoman 
manual](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview)

```
yo office --name "My Test AddIn" --host word --ts true
```

# local development

The following plugin is necessary:
- `office-addin-debugging`
- `office-addin-dev-certs`

See `package.json` settings:
- `"config": {"app_to_debug": "word", "app_type_to_debug": "desktop", "dev_server_port": 3000}`
- `"start": "office-addin-debugging start manifest.xml"`

and `webpack.config.js`'s `devServer` settings.