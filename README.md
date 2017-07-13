XlsxEditor
========

Edit xlsx cell with ruby

Installation
--------

Add this line to your application's Gemfile:

```ruby
gem 'xlsx_editor'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install xlsx_editor

Usage
--------

* Show value for cell

```ruby
XlsxEditor.fetch(
  path:       '/path/to/your/xlsx',
  sheet_name: 'Sheet1',
  cells:      %w[A1 C4 F12]
)
```

* Write value to cell

``` ruby
XlsxEditor.update(
  path:        '/path/to/your/xlsx',
  output_file: '/path/to/output/path',
  sheet_name:  'Sheet1',
  cells:       {
    'A1'  => 'hoge',
    'C4'  => 'fuga',
    'F12' => 'hoge'
  }
)
```


License
--------

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

