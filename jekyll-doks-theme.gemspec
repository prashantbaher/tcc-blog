# frozen_string_literal: true

require_relative "lib/jekyll-doks-theme/version"

Gem::Specification.new do |spec|
  spec.name          = "jekyll-doks-theme"
  spec.version       = JekyllDoksTheme::VERSION
  spec.authors       = ["Your Name"]
  spec.email         = ["you@example.com"]

  spec.summary       = "Doks-style documentation theme for Jekyll"
  spec.homepage      = "https://github.com/your-username/jekyll-doks-theme"
  spec.license       = "MIT"

  spec.files         = Dir["{_layouts,_includes,_sass,assets}/**/*", "LICENSE", "README.md", "lib/**/*"]
  spec.add_runtime_dependency "jekyll", ">= 4.0"
end
