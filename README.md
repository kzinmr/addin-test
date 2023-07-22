# contract assistant

ラフスケッチ

## drafting

### build phase

- upload template
- analyze template & generate form items
  - how to generate switching items?
  - with/without LLM
- configure form items (form label, form type, optional or not)
- save these configuration as task template

### runtime phase

- create a new task from task template configured in build step
- prompt a user to fill in the forms in the task
- fill in the placeholders in contract template based on the user input values
- generate document data state
- export .docx file from generated data as user downloadable format

### contents management

- edit template
- apply changes


## review

### office addin frontend
- search and edit
  - 1. select text
  - 2. search selected text
  - 3. show search results
  - (switch edit mode...)
  - 4. show diff between selected text and selected search result
  - 5. refine selected search result
  - 6. insert selected search result
- prompt and edit
  - 1. input prompt query
  - 2. refine selected result
  - 3. insert selected result
- simple translation
  - 1. select text
  - 2. translate (set language)
  - 3. insert result

### web app frontend
- management UI
  - show stats of documents
  - explore registered documents
  - view selected document
  - edit clauses and text in document
- simple search UI
  - input clause (free text to vectorize)

### backend
- search server
  - indexing
  - search
- documents analysis server
  - upload
  - analyze
    - party
    - tag
    - clauses & title
- llm prompt server
  - configuration
  - process
- document management app server
  - explorer
  - viewer
  - editor

## infrastructure

tbd
