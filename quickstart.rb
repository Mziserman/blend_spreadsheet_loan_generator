require 'active_support/all'
require "google/apis/sheets_v4"
require "googleauth"
require "googleauth/stores/file_token_store"
require "fileutils"

OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "Google Sheets API Ruby Quickstart".freeze
CREDENTIALS_PATH = "credentials.json".freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = "default"
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
    url = authorizer.get_authorization_url base_url: OOB_URI
    puts "Open the following URL in the browser and enter the " \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize

# Prints the names and majors of students in a sample spreadsheet:
# https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
# spreadsheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
# range = "Class Data!A2:E"
# response = service.get_spreadsheet_values spreadsheet_id, range
# puts "Name, Major:"
# puts "No data found." if response.values.empty?
# response.values.each do |row|
#   # Print columns A and E, which correspond to indices 0 and 4.
#   puts "#{row[0]}, #{row[4]}"
# end

puts "amount ?"
# amount = gets.chomp.to_f || 1000.0
amount = 1000.0

puts "duration ?"
# duration = gets.chomp.to_i  || 12
duration = 12

puts "due_on ?"
# due_on = Date.parse(gets.chomp) || Date.today
due_on = Date.today

puts "period duration ? (in months)"
# period_duration = gets.chomp.to_i || 1
period_duration = 1

puts 'rate ?'
# rate = gets.chomp.to_f || 0.12
rate = 0.12

puts 'deferred and capitalized ?'
# deferred_and_capitalized = gets.chomp.to_i || 0
deferred_and_capitalized = 0

puts 'deferred ?'
# deferred = gets.chomp.to_i || 0
deferred = 0

spreadsheet = service.create_spreadsheet(
  {
    properties: {
      title: 'test'
    }
  },
  fields: 'spreadsheetId'
)

columns = %w[
  index
  due_on
  remaining_capital_start
  remaining_capital_end
  period_calculated_interests
  delta
  accrued_delta
  amount_to_add
  period_interests
  period_capital
  total_paid_capital_end_of_period
  total_paid_interests_end_of_period
  period_total
  capitalized_interests_start
  capitalized_interests_end
  period_rate
]
def column_letter
  {
    index: 'A',
    due_on: 'B',
    remaining_capital_start: 'C',
    remaining_capital_end: 'D',
    period_calculated_interests: 'E',
    delta: 'F',
    accrued_delta: 'G',
    amount_to_add: 'H',
    period_interests: 'I',
    period_capital: 'J',
    total_paid_capital_end_of_period: 'K',
    total_paid_interests_end_of_period: 'L',
    period_total: 'M',
    capitalized_interests_start: 'N',
    capitalized_interests_end: 'O',
    period_rate: 'P'
  }
end

def column_range(column: 'A', upto: , exclude_head: true)
  start_line = exclude_head ? 2 : 1

  "#{column}#{start_line}:#{column}#{upto}"
end

def index_to_line(index:)
  index + 1 # first term is on line 2
end

def excel_float(float:)
  float.to_s.gsub('.', ',')
end

def remaining_capital_start_formula(index:, amount:)
  base = excel_float(float: amount)

  return base if index == 2

  "=#{base} - #{column_letter[:total_paid_capital_end_of_period]}#{index - 1}"
end

def remaining_capital_end_formula(index:, amount:)
  base = excel_float(float: amount)

  "=#{base} - #{column_letter[:total_paid_capital_end_of_period]}#{index}"
end

def period_calculated_interests_formula(index:)
  "=#{column_letter[:remaining_capital_start]}#{index} * #{column_letter[:period_rate]}#{index}"
end

def period_accrued_delta_formula(index:)
  return excel_float(float: 0.0) if index == 2

  "=#{column_letter[:accrued_delta]}#{index - 1} + #{column_letter[:delta]}#{index}"
end

def total_paid_capital_end_of_period_formula(index:)
  return "=#{column_letter[:period_capital]}2" if index == 2

  "=SOMME(#{column_range(column: column_letter[:period_capital], upto: index)})"
end

def total_paid_interests_end_of_period_formula(index:)
  return "=#{column_letter[:period_interests]}2" if index == 2

  "=SOMME(#{column_range(column: column_letter[:period_interests], upto: index)})"
end

def period_rate_formula(rate:, period_duration:, type: :simple)
  case type
  when :normal
    "=TAUX.NOMINAL(#{excel_float(float: rate)};#{excel_float(float: 12.0 / period_duration)})"
  when :simple
    "=#{excel_float(float: rate)} / 12,0"
  end
end

def period_interests_formula(index:, duration:, amount:)
  "=-IPMT(#{column_letter[:period_rate]}#{index}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
end


def period_capital_formula(index:, duration:, amount:)
  "=-PPMT(#{column_letter[:period_rate]}#{index}; #{column_letter[:index]}#{index}; #{duration}; #{excel_float(float: amount)})"
end

def period_total_formula(index:, duration:, amount:, type: :standard)
  if type == :standard
    "=-PMT(#{column_letter[:period_rate]}#{index}; #{duration}; #{excel_float(float: amount)})"
  else
    "=#{column_letter[:period_capital]}#{index} + #{column_letter[:period_interests]}#{index}"
  end
end

def line(index:, due_on:, amount:, period_duration:, rate:, duration:)
  term = index + 1
  [
    term,
    due_on + (term * period_duration).months,
    remaining_capital_start_formula(index: index_to_line(index: term), amount: amount),
    remaining_capital_end_formula(index: index_to_line(index: term), amount: amount),
    period_calculated_interests_formula(index: index_to_line(index: term)),
    '',
    period_accrued_delta_formula(index: index_to_line(index: term)),
    'amount_to_add',
    period_interests_formula(index: index_to_line(index: term), duration: duration, amount: amount),
    period_capital_formula(index: index_to_line(index: term), duration: duration, amount: amount),
    total_paid_capital_end_of_period_formula(index: index_to_line(index: term)),
    total_paid_interests_end_of_period_formula(index: index_to_line(index: term)),
    period_total_formula(index: index_to_line(index: term), duration: duration, amount: amount),
    'capitalized_interests_start',
    'capitalized_interests_end',
    period_rate_formula(rate: rate, period_duration: period_duration)
  ]
end

range = "A1:P#{duration + 1}"

value_range_object = Google::Apis::SheetsV4::ValueRange.new(
  range: range,
  values: [columns] +
    duration.times.map.with_index do |_, index|
      line(index: index, amount: amount, due_on: due_on, period_duration: period_duration, rate: rate, duration: duration)
    end

)
result = service.update_spreadsheet_value(
  spreadsheet.spreadsheet_id,
  range,
  value_range_object,
  value_input_option: 'USER_ENTERED'
)
puts "#{result.updated_cells} cells updated."
