# OOXML Excel

TODO: Description

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ooxml_excel'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ooxml_excel

## Usage

Using `OOXML::Excel` to read spreadsheet:
```
ooxml_excel = OOXML::Excel.new('example.xlsx')
```

Fetching all sheets:
```
ooxml_excel.sheets # ['Test Sheet 1', 'Test Sheet 2']
```

### Iterating to each row:
```
# as an array of strings
ooxml_excel.sheet('Test Sheet 1').each_row do |row|
  # Do something here...
end

# as an array of objects
ooxml_excel.sheet('Test Sheet 1').each_row_as_object do |row|
  # Do something here...
  p row
end

```

### Fetching Columns:
```
# Fetch all columns
oooxml_excel.columns

# Checking if the column is hidden
ooxml_excel.column('A1').hidden?
```

### Fetching Styles:

```
# Font
font_object = ooxml_excel.sheet('Test Sheet 1').font('A1')
font.bold? # false
font.name # Arial
font.rgb_color # FFE10000
font.size # 8

# Cell Fill
fill_object = ooxml_excel.sheet('Test Sheet 1').fill('A1')
fill_object.bg_color # FFE10000
fill_object.fg_color # FFE10000
```

### Data Validation
```
# All Validations
data_validations = ooxml_excel.sheet('Test Sheet 1').data_validations

# Specific validation for cell
data_validation = ooxml.sheet('Input Sheet').data_validation_for_cell('D4')

data_validation.prompt # "Sample Validation Message"
data_validation.formula # 20
data_validation.type #textLength

```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/ooxml_excel.
