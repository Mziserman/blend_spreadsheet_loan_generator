require_relative 'lib/blend_spreadsheet_loan_generator/version'

Gem::Specification.new do |spec|
  spec.name          = 'blend_spreadsheet_loan_generator'
  spec.version       = BlendSpreadsheetLoanGenerator::VERSION
  spec.authors       = ['MZiserman']
  spec.email         = ['martinziserman@gmail.com']

  spec.summary       = 'Generate spreadsheets amortization schedules from the command line'
  spec.homepage      = 'https://github.com/CapSens/blend_spreadsheet_loan_generator'
  spec.required_ruby_version = Gem::Requirement.new('>= 2.3.0')

  spec.metadata['allowed_push_host'] = 'https://rubygems.org'

  spec.metadata['homepage_uri'] = spec.homepage
  spec.metadata['source_code_uri'] = spec.homepage
  # spec.metadata["changelog_uri"] = "TODO: Put your gem's CHANGELOG.md URL here."

  # Specify which files should be added to the gem when it is released.
  # The `git ls-files -z` loads the files in the RubyGem that have been added into git.
  spec.files = Dir.chdir(File.expand_path('..', __FILE__)) do
    `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  end
  spec.bindir        = 'exe'
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.executables << 'bslg'

  spec.require_paths = ['lib']

  spec.add_runtime_dependency 'activesupport'
  spec.add_runtime_dependency 'csv'
  spec.add_runtime_dependency 'dry-cli', '0.6'
  spec.add_runtime_dependency 'google_drive'
  spec.add_development_dependency 'pry'
end
