# Open_IRIS_tools

To make sure you have the latest version, please follow these steps:
- Click green "Code" download button -> Download ZIP.
- Save and extract .zip file in a new folder.
- Start a new Jupyter Notebook instance and access the new folder. 

NOTE: to download a specific branch (other than 'master') you might have to use a specific URL, e.g. https://github.com/UH-LMU/Open_IRIS_tools/archive/add_products.zip to get the 'add_products' branch.

## Before committing

Notebook cell outputs can contain real booking/invoice data (names, emails, WBS codes).
This repo uses `nbstripout` to strip outputs automatically before they're committed, but
it has to be enabled once per clone:

```
pip install nbstripout
nbstripout --install
```

After that, `git add`/`git commit` will silently strip notebook outputs for you. CI also
checks for this on every push as a backstop, in case a clone doesn't have it enabled.
