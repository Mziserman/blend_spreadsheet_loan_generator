require 'dry/cli'
require 'google_drive'
require 'active_support/all'
require 'fileutils'
require 'csv'

require 'blend_spreadsheet_loan_generator/version'

module BlendSpreadsheetLoanGenerator
  class Error < StandardError; end

  extend Dry::CLI::Registry

  autoload :SpreadsheetConcern, 'blend_spreadsheet_loan_generator/concerns/spreadsheet_concern'
  autoload :CsvConcern, 'blend_spreadsheet_loan_generator/concerns/csv_concern'

  autoload :Linear, 'blend_spreadsheet_loan_generator/linear'
  autoload :Standard, 'blend_spreadsheet_loan_generator/standard'
  autoload :NormalInterests, 'blend_spreadsheet_loan_generator/normal_interests'
  autoload :SimpleInterests, 'blend_spreadsheet_loan_generator/simple_interests'
  autoload :RealisticInterests, 'blend_spreadsheet_loan_generator/realistic_interests'
  autoload :Formula, 'blend_spreadsheet_loan_generator/formula'

  autoload :Loan, 'blend_spreadsheet_loan_generator/loan'

  autoload :Version, 'blend_spreadsheet_loan_generator/version'
  autoload :Generate, 'blend_spreadsheet_loan_generator/generate'
  autoload :Restructure, 'blend_spreadsheet_loan_generator/restructure'
  autoload :Init, 'blend_spreadsheet_loan_generator/init'

  register 'init', Init, aliases: ['i', '-i', '--init']
  register 'version', Version, aliases: ['v', '-v', '--version']
  register 'generate', Generate, aliases: ['g', '-g', '--generate']
  register 'restructure', Restructure, aliases: ['r', '-r', '--restructure']
end
