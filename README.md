# OOXL

TODO: Description

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'ooxl'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install ooxl

## Usage

### reading an excel spreadsheet:
```
ooxl = OOXL.new('example.xlsx')

or

ooxl = OOXL.open('example.xlsx')
```

### Fetching all sheets:
```
ooxl.sheets # ['Test Sheet 1', 'Test Sheet 2']
```

### Accessing Rows and Cells
```
sheet = ooxl.sheet('Test Sheet 1')

# Rows
sheet.rows[0] # Access the first row
sheet[0] # short version

# Cells
sheet.rows[0].cells # access the cells of the first row
sheet[0].cells # short version

sheet.rows[0].cells[0] # Access the first cell of the row
sheet[0][0] # short version

sheet.rows[0].cells[0].value # Access cell value
sheet[0][0].value# short version

# Fetch cell value using the short versions
ooxl['Test Sheet 1'][0][0].value

# Detecting merged cell
ooxl['Test Sheet 1'].in_merged_cell?('C1') # true/false
```

### Iteration
```
ooxl.sheet('Test Sheet 1').each do |row|
  row.each_with_index do |cell, cell_index|
    # do something here..
  end
end

ooxl.each do |sheet|
  sheet.each do |row|
    row.each do |cell|
      # do something here...
    end
  end
end
```

### Fetching Columns
```
# Fetch all columns
ooxl.sheet('Test Sheet 1').columns

# Checking if the column is hidden
ooxl.sheet('Test Sheet 1').column(1).hidden? # column index
ooxl.sheet('Test Sheet 1').column('A').hidden? # column letter
```

### Fetching Styles
```
# Font
font_object = ooxl.sheet('Test Sheet 1').font('A1')
font_object.bold? # false
font_object.name # Arial
font_object.rgb_color # FFE10000
font_object.size # 8

# Cell Fill
fill_object = ooxl.sheet('Test Sheet 1').fill('A1')
fill_object.bg_color # FFE10000
fill_object.fg_color # FFE10000
```
### Fetching Data from named/cell range
```
# named range
ooxl.named_range('my_named_range') # ['value' 'from', 'range']

# cell range
ooxml['Lists'!$A$1:$A$6] # ['1','2','3','4','5','6']

# or
ooxml['Lists'!A1:A6] # ['1','2','3','4','5','6']

# or loading a single value
ooxml['Lists'!A1] # ['1']

# or loading a box type values
ooxml['Lists!A1:B2'] # [['1', '2'], ['2','3']]

```
### Fetching Data Validation
```
# All Validations
data_validations = ooxl.sheet('Test Sheet 1').data_validations

# Specific validation for cell
data_validation = ooxml.sheet('Input Sheet').data_validation('D4')

data_validation.prompt # "Sample Validation Message"
data_validation.formula # 20
data_validation.type #textLength

```

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake spec` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and tags, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/halcjames/ooxl.
