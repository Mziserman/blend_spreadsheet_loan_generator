require 'dry/cli'

module BlendSpreadsheetLoanGenerator
  VERSION = '0.1.0'

  class Version < Dry::CLI::Command
    desc 'Print version'

    def call(*)
      puts VERSION
    end
  end
end
