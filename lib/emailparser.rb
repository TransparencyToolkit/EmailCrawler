require 'pry'
require 'json'
require 'mail'

class Emailparser

	def initialize(message, attachment_dir)
		@message = message
		@attachment_dir = attachment_dir
	end

	# Voodoo to fix nasty encoded strings
	def fix_encode(str)
		if str.is_a?(String)
		  return str.unpack('C*').pack('U*')
		else
		  return str
		end
	end

	# Accepts a message
	def parse_message

		print "loading file: " + @message

		email = Mail.read(@message)

		# Defaults
		source_file = @message.split("/").last

		if email.subject
			subject = fix_encode(email.subject)
		else 
			subject = "No Subject"
		end
		
		body_plain	= ""
		body_html	= ""
		attachments = []

		# Check for Multipart
		if email.multipart?

			# Parse Parts
			email.parts.map do |part|
				if (part.content_type.start_with?('text/plain'))
					body_plain = fix_encode(part.body.decoded)
				elsif (part.content_type.start_with?('text/html'))
					body_html = fix_encode(part.body.decoded)
				elsif (part.content_type.start_with?('multipart/alternative'))
					print "has multipart/alternative\n"
				else
					print "has part: " + part.content_type + "\n"
				end
				print "-------------------------------------------------------\n"
			end

			body_plain = fix_encode(email.text_part.body.decoded)
			body_html = fix_encode(email.html_part.body.decoded)

		else
			print "found single part email\n"
			body_plain = fix_encode(email.body.decoded)
			body_html = fix_encode(email.body.decoded)
		end

		# Handle Attachments
		email.attachments.each do | attachment |
			if (attachment.content_type.start_with?('image/'))
				filename = fix_encode(attachment.filename)
				attachments.push(filename)
				print "found attachment " + filename + "\n"
				#begin
				#	File.open(@attachment_dir + filename, "w+b", 0644) do |f|
				#		f.write attachment.body.decoded 
				#	end
				#rescue => e
				#	puts "Unable to save data for #{filename} because #{e.message}"
				#end
			end
		end

		# Structure Data
		email_data = {
			source_file: source_file,
			message_id: email.message_id,
			from: email.from,
			to: email.to,
			cc: email.cc,
			subject: subject,
			body_plain: body_plain,
			body_html: body_html,
			attachments: attachments
		}

		email_json = JSON.pretty_generate(email_data)
		return email_json
	end

end
