from tornado import ioloop, web, httpclient
import facebook
import json
import xlsxwriter

print('restarted!')

fb_secret = ''
fb_id = ''
fb_url = 'http://678664de.ngrok.com/' #test url


class MainHandler(web.RequestHandler):
    def get(self):

        self.write("<a href='https://www.facebook.com/dialog/oauth?client_id="+fb_id+"&redirect_uri="+fb_url+"success&scope=read_stream'>Click to download your links</a> (Can take a while)")

class FBHandler(web.RequestHandler):
    def get(self):
        print('here')
        code = self.get_argument("code", None, True)
        if code == None:
            error = self.get_argument("error_message", None, True)
            if error != None:
                self.write('error: ' + error)
            else:
                self.write('nope (missing fb code)')
            self.finish()

        client = httpclient.HTTPClient()
        resp = client.fetch("https://graph.facebook.com/oauth/access_token?client_id="+fb_id+"&redirect_uri="+fb_url+"success&client_secret="+fb_secret+"&code="+str(code))

        token = resp.body.replace('access_token=','')
        pos = token.find('&expires')
        token = token[:pos]

        graph = facebook.GraphAPI(token)
        profile = graph.get_object('me')

        user = profile['id']

        workbook = xlsxwriter.Workbook('excel/'+user+'.xlsx')
        worksheet = workbook.add_worksheet()

        worksheet.write(0,0,"ID")
        worksheet.write(0,1,"DateTime")
        worksheet.write(0,2,"Text")
        worksheet.write(0,3,"Link")

        #date_format = workbook.add_format({'num_format': 'mmm d yyyy hh:mm AM/PM'})

        row = 1

        feed = graph.get_object("me/links")
        while 'next' in feed['paging']:
            for post in feed['data']:
                if 'message' not in post:
                    post['message'] = 'none'

                worksheet.write(row,0,str(row))
                worksheet.write(row,1,str(post['created_time']))
                try:
                    worksheet.write(row,2,str(post['message']))
                except Exception:
                    worksheet.write(row,2,'error')

                try:
                    worksheet.write(row,3,str(post['link']))
                except Exception:
                    worksheet.write(row,3,'error')

                row += 1


            response = client.fetch(feed['paging']['next'])
            feed = json.loads(str(response.body))

        workbook.close()

        self.write("<a href='download/"+user+".xlsx'>Download Spreadsheet.</a>")

class DownloadHandler(web.RequestHandler):
    def get(self):
        id = self.get_argument("id", None, True)
        if id == None:
            self.write('nope')
            self.finish()






application = web.Application([
    (r"/", MainHandler),
    (r"/success", FBHandler),
    (r"/success([^/]*)", FBHandler),
    (r"/download",web.StaticFileHandler, {'path': 'excel'}),
    (r"/download/", web.StaticFileHandler, {'path': 'excel'}),
    (r"/download/([^/]*)", web.StaticFileHandler, {'path': 'excel'}),
], debug = True)

if __name__ == "__main__":
    application.listen(8008)
    ioloop.IOLoop.instance().start()
