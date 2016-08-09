require 'pry'
require 'json'
require 'mail'

class Emailparser

	def initialize(message, out_dir, attachment_dir)
		@message = message
		@attachment_dir = out_dir + "/" + attachment_dir
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

		puts "loading email: " + @message + "\n"

		email = Mail.read(@message)

		# Defaults
		source_file = @message.split("/").last

		# Addresses
		begin
			email_to = email.to.to_a
		rescue
			email_to = [email.to.gsub("<", "").gsub(">", "")]
		end
		begin
			email_cc = email.cc.to_a
		rescue
			email_cc = [email.cc.gsub("<", "").gsub(">", "")]
		end
		begin
			email_from = email.from.to_a
		rescue
			email_from = [email.from.gsub("<", "").gsub(">", "")]
		end
		begin
			recipients = email_to.concat(email_cc)
			addresses = recipients + email_from
		rescue
			puts "oops something failed here..."
			# binding.pry
		end

		# Subject
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
			puts "- is a multipart email\n"
			if email.text_part 
				body_plain = fix_encode(email.text_part.body.decoded)
			end
			if email.html_part
				body_html = fix_encode(email.html_part.body.decoded)
			end
		else
			puts "- is single part email\n"
			body_plain = fix_encode(email.body.decoded)
			body_html = fix_encode(email.body.decoded)
		end

		# Handle Attachments
		email.attachments.each do | attachment |
			if (attachment.content_type.start_with?('image/'))
				filename = fix_encode(attachment.filename)
				attachments.push(filename)
				print "- found attachment " + filename + "\n"
				begin
					File.open(@attachment_dir + filename, "w+b", 0644) do |f|
						f.write attachment.body.decoded 
					end
				rescue => e
					puts "Unable to save data for #{filename} because #{e.message}"
				end
			end
		end

		# Structure Data
		email_data = {
			source_file: source_file,
			message_id: email.message_id,
			date: email.date,
			sender: email.from,
			from: email.from,
			to: email.to,
			cc: email.cc,
			recipients: recipients,
			addresses: addresses,
			subject: subject,
			body_plain: body_plain,
			body_html: body_html,
			attachments: attachments
		}

		email_json = JSON.pretty_generate(email_data)
		return email_json
	end

end
