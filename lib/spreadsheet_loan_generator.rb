require 'bundler/setup'
require 'dry/cli'
require 'google_drive'
require 'active_support/all'
require 'fileutils'
require 'csv'

require 'spreadsheet_loan_generator/version'

module SpreadsheetLoanGenerator
  class Error < StandardError; end

  extend Dry::CLI::Registry

  autoload :SpreadsheetConcern, 'spreadsheet_loan_generator/concerns/spreadsheet_concern'
  autoload :CsvConcern, 'spreadsheet_loan_generator/concerns/csv_concern'
  autoload :Loan, 'spreadsheet_loan_generator/loan'

  autoload :Formula, 'spreadsheet_loan_generator/formula'
  autoload :Linear, 'spreadsheet_loan_generator/formula/linear'
  autoload :Standard, 'spreadsheet_loan_generator/formula/standard'
  autoload :NormalInterests, 'spreadsheet_loan_generator/formula/normal_interests'
  autoload :SimpleInterests, 'spreadsheet_loan_generator/formula/simple_interests'

  autoload :Version, 'spreadsheet_loan_generator/version'
  autoload :Generate, 'spreadsheet_loan_generator/generate'
  autoload :Init, 'spreadsheet_loan_generator/init'

  register 'init', Init, aliases: ['i', '-i', '--init']
  register 'version', Version, aliases: ['v', '-v', '--version']
  register 'generate', Generate, aliases: ['g', '-g', '--generate']
end
