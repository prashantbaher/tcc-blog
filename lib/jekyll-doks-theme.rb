# frozen_string_literal: true

require "jekyll-doks-theme/version"

module JekyllDoksTheme
  class << self
    def load!
      Jekyll::Hooks.register :site, :after_init do |site|
        # nothing special yet
      end
    end
  end
end

JekyllDoksTheme.load!
