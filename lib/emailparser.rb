require 'pry'
require 'json'
require 'mail'
require 'digest'

class Emailparser

	def initialize(message, out_dir, attachment_dir)
		@message = message
		@attachment_dir = out_dir + "/" + attachment_dir
		@allowed_documents = [
			'application/x-mobipocket-ebook',
			'application/epub+zip',
			'application/rtf',
			'application/vnd.ms-works',
			'application/msword',
			'application/pdf',
			'application/x-download',
			'message/rfc822',
			'text/x-log',
			'text/scriptlet',
			'text/plain',
			'text/iuls',
			'text/plain',
			'text/richtext',
			'text/x-setext',
			'text/x-component',
			'text/webviewhtml',
			'text/h323',
			'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			'application/vnd.oasis.opendocument.text',
			'application/vnd.oasis.opendocument.text-template',
			'application/vnd.sun.xml.writer',
			'application/vnd.sun.xml.writer.template',
			'application/vnd.sun.xml.writer.global',
			'application/vnd.stardivision.writer',
			'application/vnd.stardivision.writer-global',
			'application/x-starwriter'
		]
		@allowed_spreadsheets = [
			'application/excel',
			'application/msexcel',
			'application/vnd.ms-excel',
			'application/vnd.msexcel',
			'application/csv',
			'application/x-csv',
			'text/tab-separated-values',
			'text/x-comma-separated-values',
			'text/comma-separated-values',
			'text/csv',
			'text/x-csv',
			'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			'application/vnd.oasis.opendocument.spreadsheet',
			'application/vnd.oasis.opendocument.spreadsheet-template',
			'application/vnd.sun.xml.calc',
			'application/vnd.sun.xml.calc.template',
			'application/vnd.stardivision.calc',
			'application/x-starcalc'
		]
		@allowed_audio = [
			'audio/amr',
			'audio/mp3',
			'audio/midi',
			'audio/mid',
			'audio/mpeg',
			'audio/basic',
			'audio/x-aiff',
			'audio/x-pn-realaudio',
			'audio/x-pn-realaudio',
			'audio/mid',
			'audio/basic',
			'audio/x-wav',
			'audio/x-mpegurl',
			'audio/wave',
			'audio/wav',
			'audio/mp4a-latm'
		]
		@allowed_contacts = [
			'text/directory',
			'text/x-vcard',
			'text/x-ms-contact'
		]
		@allowed_images = [
			'image/png',
			'image/jpeg',
			'image/cis-cod',
			'image/ief',
			'image/pipeg',
			'image/tiff',
			'image/x-cmx',
			'image/x-cmu-raster',
			'image/x-rgb',
			'image/x-icon',
			'image/x-xbitmap',
			'image/x-xpixmap',
			'image/x-xwindowdump',
			'image/x-portable-anymap',
			'image/x-portable-graymap',
			'image/x-portable-pixmap',
			'image/x-portable-bitmap',
			'image/svg+xml',
			'application/x-photoshop',
			'application/postscript'
		]
		@allowed_slideshows = [
			'application/powerpoint',
			'application/vnd.ms-powerpoint',
			'application/vnd.oasis.opendocument.presentation',
			'application/vnd.oasis.opendocument.presentation-template',
			'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			'application/vnd.sun.xml.impress',
			'application/vnd.sun.xml.impress.template',
			'application/vnd.stardivision.impress',
			'application/vnd.stardivision.impress-packed',
			'application/x-starimpress'
		]
		@allowed_videos = [
			'video/quicktime',
			'video/x-sgi-movie',
			'video/mpeg',
			'video/x-la-asf',
			'video/x-ms-asf',
			'video/x-msvideo',
			'video/mp4',
			'video/mp2',
			'video/avi'
		]
	end

	# Voodoo to fix nasty encoded strings
	def fix_encode(str)
		if str.is_a?(String)
		  return str.unpack('C*').pack('U*')
		else
		  return str
		end
	end

	def make_attachment_folder(attachments, source_hash)
		if (!attachments.empty?)
			puts "Creating attachments directory: " + source_hash
			attachments_dir = @attachment_dir + source_hash
      		Dir.mkdir(attachments_dir) if !Dir.exist?(attachments_dir)
		end
	end

	def save_attachment(attachment, message_id, filename)
		puts " - found attachment " + filename + "\n"
		begin
			File.open(@attachment_dir + message_id + "/" + filename, "w+b", 0644) do |f|
				f.write attachment.body.decoded 
			end
		rescue => e
			puts "Unable to save data for #{filename} because #{e.message}"
		end
	end

	# Accepts a message
	def parse_message

		puts "Loading email: " + @message + "\n"
		source_hash = Digest::MD5.hexdigest(File.read(@message))
		puts "Hash of email: " + source_hash

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
			puts " - is multipart\n"
			if email.text_part 
				body_plain = fix_encode(email.text_part.body.decoded)
			end
			if email.html_part
				body_html = fix_encode(email.html_part.body.decoded)
			end
		else
			puts " - is single part\n"
			body_plain = fix_encode(email.body.decoded)
			body_html = fix_encode(email.body.decoded)
		end

		# Handle Attachments
		make_attachment_folder(email.attachments, source_hash)
		email.attachments.each do | attachment |
			attachment_save = false
			filename = fix_encode(attachment.filename)
			mime_type, remaining = attachment.content_type.split(';', 2)
			puts " - Attachment mime: " + mime_type
			# Check Allowed Mime Types
			if (@allowed_documents.include? mime_type)
				puts " - Attachment is: document"
				attachment_save = true
			elsif (@allowed_spreadsheets.include? mime_type)
				puts " - Attachment is: spreadsheet"
				attachment_save = true
			elsif (@allowed_audio.include? mime_type)
				puts " - Attachment is: audio"
				attachment_save = true
			elsif (@allowed_contacts.include? mime_type)
				puts " - Attachment is: contact"
				attachment_save = true
			elsif (@allowed_images.include? mime_type)
				puts " - Attachment is: image"
				attachment_save = true
			elsif (@allowed_slideshows.include? mime_type)
				puts " - Attachment is: slideshow"
				attachment_save = true
			elsif (@allowed_videos.include? mime_type)
				puts " - Attachment is: video"
				attachment_save = true
			end

			# Process Attachment
			if (attachment_save == true)	
				attachments.push(source_hash + "/" + filename)
				save_attachment(attachment, source_hash, filename)
			end
		end

		# Structure Data
		email_data = {
			source_file: source_file,
			source_hash: source_hash,
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
