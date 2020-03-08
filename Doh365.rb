#!/usr/bin/env ruby

require "mechanize" # v1 only
require "net/http"
require "uri"
require "json"
require "colorize"

class UhOh365NG
  attr_accessor :wait
  attr_reader   :agent

  def initialize(wait:, agent: nil)
    @agent = agent
    @wait  = wait
  end

  def verify_v1(email)
    begin
      code = agent.get("https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/#{email}?Protocol=Autodiscoverv1").code
      return [email, code.to_i]
    rescue Exception
      return [email, 302]
    end
  end

  def verify_v2(email)
      uri = URI.parse("https://login.microsoftonline.com/common/GetCredentialType")
      request = Net::HTTP::Post.new(uri)
      request.body = JSON.dump({
        "username"             => email,
        "isOtherIdpSupported"  => true,
        "isRemoteNGCSupported" => true,
        "isFidoSupported"      => true,
        "checkPhones"          => false,
        "isCookieBannerShown"  => false,
      })

      begin
        response = Net::HTTP.start(uri.hostname, uri.port, use_ssl: true) do |http|
          http.request(request)
        end
        body = JSON.parse(response.body)
      rescue Exception => e
        puts e
        return [false, nil]
      end

      is_throttled  = body["ThrottleStatus"] != 0
      result_exists = body["IfExistsResult"] != 1
      is_valid      = result_exists && !is_throttled

      if is_throttled
        puts "[-] We're being throttled now! Waiting and increasing wait by 1.7x"
        wait = 1.7 * wait
        puts "Wait set to #{wait}"
        puts "Sleeping #{wait * 10}"
        sleep(wait * 8)
      end

      [is_valid, is_throttled]
  end
end

# For use with v1 check
# agent = Mechanize.new do |agent|
#   agent.open_timeout   = 30
#   agent.read_timeout   = 30
#   agent.request_headers = {"Accept" => "application/json"}
#   agent.user_agent = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0.12026; Pro)"
# end

ARGV.length == 1 || raise("Requires new-line seperated list of emails as only arg")

verifier = UhOh365NG.new(wait: 0.8)
emails   = File.readlines(ARGV[0]).map(&:chomp)
start    = Time.now

results = emails.each_with_index.with_object([]) do |(email, idx), memo|
  t_start = Time.now
  valid, throttled = verifier.verify_v2(email)
  entry = {
    "id"        => idx + 1,
    "email"     => email,
    "valid"     => valid,
    "throttled" => throttled,
    "time"      => Time.now - t_start,
  }

  puts valid ? entry.to_s.blue : entry.to_s.red
  if throttled
    sleep(verifier.wait)
    redo
  end
  memo.push(entry) if valid
end
total_time = Time.now - start

puts("Discovered #{results.length} valid emails")
File.open("save.json", "w+") { |file| file.write(JSON.pretty_generate(results)) }
puts("Completed in #{total_time}")
