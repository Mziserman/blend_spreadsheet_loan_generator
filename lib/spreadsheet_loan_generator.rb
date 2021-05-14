require 'bundler/setup'
require 'dry/cli'
require 'google_drive'
require 'active_support/all'
require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'fileutils'


require 'spreadsheet_loan_generator/version'

module SpreadsheetLoanGenerator
  class Error < StandardError; end

  extend Dry::CLI::Registry

  autoload :SpreadsheetConcern, 'spreadsheet_loan_generator/concerns/spreadsheet_concern'
  autoload :FormulaConcern, 'spreadsheet_loan_generator/concerns/formula_concern'
  autoload :ServiceWrapper, 'spreadsheet_loan_generator/service_wrapper'
  autoload :Loan, 'spreadsheet_loan_generator/loan'

  autoload :Version, 'spreadsheet_loan_generator/version'
  autoload :Generate, 'spreadsheet_loan_generator/generate'

  register 'version', Version, aliases: ['v', '-v', '--version']
  register 'generate', Generate, aliases: ['g', '-g', '--generate']
end
