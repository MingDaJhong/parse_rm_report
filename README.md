# ParseName

## Requirement

1. node.js
2. yarn

## Version
- node.js: available v14, v16, v18

## Installation
### Install yarn

```
# install (ios)
brew install yarn

# install (npm)
npm install --global yarn

# check version
yarn --version

# check yarn path
which yarn
```

## Usage / Example

### Install dependencies
```
yarn install
```
- **Only accept .xlsx !!!**

```
node main filePath
```
**Example :**

```
node main testing.xlsx
```

## Dependencies
- https://github.com/dhuddell/convert-xls-to-json
- https://github.com/LuisEnMarroquin/json-as-xlsx
