require 'dry/cli'

module SpreadsheetLoanGenerator
  VERSION = '0.1.4'

  class Version < Dry::CLI::Command
    desc 'Print version'

    def call(*)
      puts VERSION
    end
  end
end
