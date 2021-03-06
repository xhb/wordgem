# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'word_handler/version'

Gem::Specification.new do |spec|
  spec.name          = "word_handler"
  spec.version       = WordHandler::VERSION
  spec.authors       = ["xhb"]
  spec.email         = ["progamstart@163.com"]
  spec.summary       = %q{provide a WordHandler for dtp vm test}
  spec.description   = %q{provide a WordHandler for dtp vm test}
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0")
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.5"
  spec.add_development_dependency "rake"
  
end
