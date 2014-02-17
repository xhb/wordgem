# WordHandler

1.封装在windows系统下操作word的com组件接口，让在windows下使用ruby读写word更加容易。
2.提供一个命令行工具，wpsword ，
  1.使用该工具批量产生和删除word文档。
  2.批量读写word文档


## Installation

Add this line to your application's Gemfile:

    gem 'word_handler'

And then execute:

    $ bundle 

Or install it yourself as:

    $ gem install word_handler

## Usage

1.在ruby脚本中使用
  require  "word_handler"
  
  word = WPS::Word.new()
  
  word.givedoc("xhb_doc", "D:\\xhb_doc")
  word.givedoc("cry_doc", "D:\\cyr_doc")
  
  word.msg("xhb_doc", "hello, here is my msg 1")
  word.msg("cry_doc", "hello, here is my msg 111")
  word.msg("xhb_doc", "hello, here is my msg 2")
  word.msg("cry_doc", "hello, here is my msg 222")

  word.closedoc("xhb_doc")
  word.closedoc("cry_doc")
  
  word.close

2.命令行方式使用
  
  wpsword  -v -f "D:\\1.wps"  -m "here is my first msg from commond line" 

## Contributing

1. Fork it ( http://github.com/<my-github-username>/word_handler/fork )
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
