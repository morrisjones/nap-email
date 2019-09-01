

import os
import argparse
from smtplib import SMTP
import csv
import re
import json
from __init__ import __version__
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from nap.prereg import Seat
from nap.gamefile.player import canonical_pnum


class EmailUF(object):

  def __init__(self):
    self.smtphost = os.environ.get('SMTP_HOST')
    self.smtpuser = os.environ.get('SMTP_USER')
    self.smtppass = os.environ.get('SMTP_PASSWORD')
    self.prereg_dir = os.environ.get('PREREG_DIR')
    self.unitfinal = None
    self.player_info = os.environ.get('PLAYER_INFO')
    self.mergefile = None
    self.test = False
    self.test_email_to = 'mojo@whiteoaks.com'
    self.verbose = False
    self.email_reference = {}
    self.nap = None
    return

  def main(self,arglist):

    parser = argparse.ArgumentParser(description="""Custom NAP qualifier email.
    Works in two phases, generate a json merge file (--json), then merge the file
    (--mergefile) with an email template to send emails.
    The email templates and images are hard coded at the moment:
      html: nap-template.html
      text: nap-template.txt
      image: images/nap-w500-h269.png
    """)
    parser.add_argument('-V', '--version', action="version",
        version="%(prog)s ("+__version__+")")
    parser.add_argument('-v', '--verbose', action="store_true", help='Verbose messages')
    parser.add_argument('-s', '--smtphost',
                        help='STMP mail host name, overrides $SMTP_HOST')
    parser.add_argument('-u', '--user',
                        help='User for SMTP login, overrides $SMTP_USER')
    parser.add_argument('-p', '--password',
                        help='Password for SMTP login, overrides $SMTP_PASSWORD')
    parser.add_argument('-t', '--test',
                        help="TEST mode, all email goes to $SMTP_TEST")
    parser.add_argument('--testconn', action="store_true", help="Test SMTP connection")
    parser.add_argument('-j','--json', action="store_true", help="Generate JSON merge data")
    parser.add_argument('-m','--mergefile', help="Specify merge file name")
    parser.add_argument('-i','--playerinfo', help="File name for email CSV data, overrides $PLAYER_INFO")
    parser.add_argument('-P','--prereg', help="Directory where UF files are found")
    parser.add_argument('--unitfinal', help="Values 1 or 2 for Unit Final game")
    args = parser.parse_args(arglist)

    errors = []

    if args.verbose:
      self.verbose = True

    #
    # These arguments required to create the JSON merge file
    #

    if args.json:
      if args.prereg:
        self.prereg_dir = args.prereg
      if not self.prereg_dir:
        errors.append("No prereg directory specified")
      if args.unitfinal:
        self.unitfinal = args.unitfinal
        if self.unitfinal not in ['1', '2']:
          errors.append("unitfinal value must be '1' or '2'")
      if not self.unitfinal:
        errors.append("Must specify which unit final game")
      if args.playerinfo:
        self.player_info = args.playerinfo
      if not self.player_info:
        errors.append("No player info file specified")
      if not errors and self.verbose:
        print "# Player info file: %s" % self.player_info

    #
    # These arguments required to generate emails from mergefile and template
    #

    if args.mergefile:

      self.mergefile = args.mergefile

      if args.smtphost:
        self.smtphost = args.smtphost
      if not self.smtphost:
        errors.append("No SMTP_HOST specified")
      if not errors and self.verbose:
        print "SMTP mail host: %s" % self.smtphost

      if args.user:
        self.smtpuser = args.user
      if not self.smtpuser:
        errors.append("No SMTP user defined")

      if not errors and self.verbose:
        print "SMTP user: %s" % self.smtpuser

      if args.password:
        self.smtppass = args.password
      if not self.smtppass:
        errors.append("No SMTP password defined")

      if not errors and self.verbose:
        print "SMTP password defined"

    #
    # Generate a JSON merge file of recipients for later
    #

    if not errors and args.json:

      if self.verbose:
        print "# Loading email info"
      self.load_d23_emails(self.player_info)
      if self.verbose:
        print "# JSON output merge data follows"
      merge = self.load_prereg()
      doc = {
        '_number_of_records': len(merge),
        'merge': merge,
      }
      print json.dumps(doc, sort_keys=True, indent=4, separators=(',',': '))

    #
    # Generate emails from mergefile plus two templates, text and html
    #
    if args.mergefile:
      with open(args.mergefile,'r') as f:
        mergefile = json.load(f)
      mergelist = mergefile['merge']

      if args.test:
        self.test = True
        self.test_email_to = args.test
        if args.verbose:
          print "Test mode, all email to %s" % self.test_email_to

      for merge in mergelist:
        partnership = merge['partnership']
        pa = partnership['player_a']
        pb = partnership['player_b']
        flight = partnership['flight']
        table = partnership['table']
        direction = partnership['direction']
        pa_name = "%s %s" % (pa['fname'],pa['lname'])
        pb_name = "%s %s" % (pb['fname'],pb['lname'])
        stationary = "Yes" if partnership['req_ns'] else "No"

        fields = {
          'pa': pa_name,
          'pb': pb_name,
          'flight': flight,
          'table': table,
          'direction': direction,
          'stationary': stationary,
        }

        # The mime object gets the To: address that appears in the message
        print "Message To: %s" % merge['email']
        msg = self.make_mime(merge['email'],fields)

        # The sendmail object gets the To: address of the envelope, where the mail is
        # actually delivered
        if args.test:
          mailto = args.test
        else:
          mailto = merge['email']
        print "actually mailing to: %s" % (mailto)
        self.sendmail("nap@bridgemojo.com",mailto,msg)


    if args.testconn:
      self.testconn()

    if errors:
      print "*** Errors:"
      for error in errors:
        print error
      return 1

    return 0

  def load_prereg(self):
    d = self.prereg_dir
    g = self.unitfinal
    merge = []
    for flight in ['a','b','c']:
      uff = "%s/UF%s-%s.json" % (d,g,flight)
      if self.verbose:
        print "# Loading %s" % uff
      with open(uff,'r') as f:
        ufsection = json.load(f)
      section = ufsection['section']
      for table in section.keys():
        for direction in (Seat.NS,Seat.EW):
          if section[table][direction]['seat']:
            a = section[table][direction]['seat']['player_a']
            b = section[table][direction]['seat']['player_b']
            partnership = {
              'player_a': {
                'pnum': a['pnum'],
                'fname': a['fname'],
                'lname': a['lname'],
              },
              'player_b': {
                'pnum': b['pnum'],
                'fname': b['fname'],
                'lname': b['lname'],
              },
              'flight': flight.upper(),
              'table': table,
              'direction': direction,
              'req_ns': section[table][direction]['seat']['req_ns']
            }
            for player in (a,b):
              cpnum = canonical_pnum(player['pnum'])
              if cpnum in self.email_reference:
                email = self.email_reference[cpnum]
                email['partnership'] = partnership
                merge.append(email)
    return merge

  def sendmail(self,mailfrom,mailto,message):
    smtp = SMTP(self.smtphost)
    # smtp.set_debuglevel(True)
    smtp.starttls()
    smtp.login(self.smtpuser,self.smtppass)
    smtp.sendmail(mailfrom,mailto,message)
    smtp.quit()

    return

  def testconn(self):
    """Test SMTP server connection"""
    smtp = SMTP(self.smtphost)
    smtp.set_debuglevel(True)
    smtp.starttls()
    smtp.login(self.smtpuser,self.smtppass)
    smtp.quit()

    return

  def load_d23_emails(self,player_info_file):
    valid_pnum = re.compile('^[0-9J-R]\d\d\d\d\d\d$')
    with open(player_info_file,'r') as f:
      csvfile = csv.reader(f)
      for row in csvfile:
        pnum, fname, lname, email = row
        if valid_pnum.match(pnum) \
            and not email.startswith('Confidential') \
            and len(email) > 0:
          canon_pnum = canonical_pnum(pnum)
          self.email_reference[canon_pnum] = {
            'pnum': pnum,
            'fname': fname,
            'lname': lname,
            'email': email,
          }
    return

  def make_mime(self, to_address, fields):

    msgRoot = MIMEMultipart('alternative')
    msgRoot['From'] = "North American Pairs District 23 <nap@bridgemojo.com>"
    msgRoot['To'] = to_address
    msgRoot['Subject'] = "Unit Final Information"

    msgHtml = MIMEMultipart('related')

    with open('uf2-template.html','r') as f:
      htmltemplate = f.read()

    html = htmltemplate.format(**fields)

    htmlpart = MIMEText(html,'html')
    msgHtml.attach(htmlpart)

    with open('uf2-template.txt','r') as f:
      texttemplate = f.read()
    text = texttemplate.format(**fields)

    textpart = MIMEText(text,'plain')
    msgRoot.attach(textpart)

    msgRoot.attach(msgHtml)

    return msgRoot.as_string()
