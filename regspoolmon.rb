require 'net/http'
require 'open-uri'
require 'win32ole'
require 'logger'

class RegSpoolMon
	$log = Logger.new('d:/regspoolmon.log')

	def map_drive
		@net = WIN32OLE.new('WScript.Network')
		@user_name = 'administrator'
		@password = 'topsecret-password'
		@net.MapNetworkDrive('v:', '\\\\coldfusion-server\\Spool$', nil, @user_name, @password)
		$log.info("Mapped network drive...")			
	end
	def disconnect_drive
		@net.RemoveNetworkDrive('v:', true)
		$log.info("Disconnected network drive.")
	end
	def get_spool_job_count
		spool_folder = 'v:'
		spoolJobCount = Dir[spool_folder + "/**/*"].length
		$log.info("Number of files in spool folder: #{spoolJobCount}")
		return spoolJobCount
	end
	def resetSpooler
			open("http://coldfusion-server/mailreset.cfm")
			$log.info("Reset ColdFusion mail spooler.")
			sleep(3)
	end
	# 	def send_email(to,opts={})
# 		  opts[:server]      ||= 'smtp-gateway'
# 		  opts[:from]        ||= 'regmonscript@domain.local'
# 		  opts[:from_alias]  ||= 'Example Emailer'
# 		  opts[:subject]     ||= "You need to see this"
# 		  opts[:body]        ||= "Important stuff!"

# 		  msg = <<END_OF_MESSAGE
# From: #{opts[:from_alias]} <#{opts[:from]}>
# To: <#{to}>
# Subject: #{opts[:subject]}

# #{opts[:body]}
# END_OF_MESSAGE

# 		  Net::SMTP.start(opts[:server]) do |smtp|
# 		    smtp.send_message msg, opts[:from], to
# 		  end
# 	end	
end

mon = RegSpoolMon.new

mon.map_drive
if mon.get_spool_job_count > 10
	mon.resetSpooler
end
# puts mon.get_spool_job_count
# sleep(5)
mon.disconnect_drive
