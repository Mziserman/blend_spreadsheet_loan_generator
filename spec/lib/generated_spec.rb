RSpec.describe 'Generated data' do
  Dir["spec/fixtures/truth/**/*.csv"].each do |fixture|
    it fixture do
      corresponding_spec = fixture.gsub('truth', 'to_test')

      expect(FileUtils.identical?(File.open(fixture), File.open(corresponding_spec))).to be true
    end
  end
end

