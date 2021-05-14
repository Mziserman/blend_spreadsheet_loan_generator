module SpreadsheetLoanGenerator
  class Init < Dry::CLI::Command

    argument :client_id, type: :string, required: true, desc: 'amount borrowed'
    argument :client_secret, type: :string, required: true, desc: 'number of reimbursements'

    def call(client_id:, client_secret:)
      if !ENV.key?('SPREADSHEET_LOAN_GENERATOR_DIR')
        puts "Please set up the environment variable SPREADSHEET_LOAN_GENERATOR_DIR to a dir that will store this gem's configuration"
        return
      end

      create_config_dir
      create_config_file(client_id: client_id, client_secret: client_secret)
    end

    def config_path
      File.join(ENV['SPREADSHEET_LOAN_GENERATOR_DIR'], 'config.json')
    end

    def create_config_dir
      if !File.exists?(ENV['SPREADSHEET_LOAN_GENERATOR_DIR'])
        FileUtils.mkdir_p(ENV['SPREADSHEET_LOAN_GENERATOR_DIR'])
      end
    end

    def create_config_file(client_id:, client_secret:)
      f = File.new(config_path, 'w')
      f.write({
        client_id: client_id,
        client_secret: client_secret
      }.to_json)
      f.close
    end
  end
end
