import random
import re
from datetime import datetime
from decimal import Decimal, getcontext
from string import digits, ascii_uppercase


decCtx = getcontext()
decCtx.prec = 7  # 5.2 digits, max=99999.99
charset = digits + ascii_uppercase

paramRE = re.compile(r"\${(\w*?)}")

def randomId(length):
    r1 = random.choice(ascii_uppercase)  # first a letter
    r2 = [random.choice(charset) for _ in range(length - 1)]  # then any mixture of capitalletters and numbers
    return r1 + ''.join(r2)

def expTitel(tour):
    return tour.getTitel()

def expEventItemId(tour):
    return tour.getEventItemId()

def expKurz(tour):
    return tour.getKurzbeschreibung()

def expFrontendLink(tour):
    return tour.getFrontendLink()

def expPublishDate(tour):
    return tour.getPublishDate().replace("T", " ");

def expDate(tour):
    t = tour.getDatum();
    return t[0][4:]

def expTime(tour):
    t = tour.getDatum();
    return t[1]

def expDuration(tour):
    b = tour.getDatumRaw()
    b = b[0:19]  # '2018-04-29T06:30:00'
    b = datetime.strptime(b, "%Y-%m-%dT%H:%M:%S")
    e = tour.getEndDatumRaw()
    e = e[0:19]  # '2018-04-29T07:30:00'
    e = datetime.strptime(e, "%Y-%m-%dT%H:%M:%S")
    d = e - b # a timedelta!
    d = str(d)[:-3] # strip :seconds
    return d

def expImageUrl(tour):
    return tour.getImageUrl()

def expOrtsname(tour):
    t = tour.getStartpunkt()
    if t[0] != "":
        return t[0] # name
    return t[1] # city

def expCity(tour):
    t = tour.getStartpunkt()
    return t[1] # city

def expStreet(tour):
    t = tour.getStartpunkt()
    return t[2] # street

def expLatitude(tour):
    t = tour.getStartpunkt()
    return str(t[3]) # latitude

def expLongitude(tour):
    t = tour.getStartpunkt()
    return str(t[4]) # longitude

def expPricing(tour):
    (minPrice, maxPrice) = tour.getPrices()
    if maxPrice == 0.0:
        return ["           <freeOfCharge>false</freeOfCharge>\n"]
    else:  return [
        f"        <fromPrice>€{minPrice}</fromPrice>\n",
        f"        <toPrice>€{maxPrice}</toPrice>\n",
        ]

def expCategories(tour):
    return ["       ???CATEGORIES???\n"]


class VADBHandler:
    def __init__(self):
        self.expFunctions = {  # keys in lower case
            "titel": expTitel,
            "eventItemId": expEventItemId,
            "publishDate": expPublishDate,
            "kurz": expKurz,
            "frontendLink": expFrontendLink,
            "date": expDate,
            "time": expTime,
            "duration": expDuration,
            "imageUrl": expImageUrl,
            "ortsname": expOrtsname,
            "latitude": expLatitude,
            "longitude": expLongitude,
            "street": expStreet,
            "city": expCity,
            "<ExpandCategories/>": expCategories,
            "<ExpandPricing/>": expPricing,
        }

        self.xmlFile = "./events.xml"
        self.outputFile = "./output.xml"
        self.output = open(self.outputFile, "w", encoding="utf-8")
        pass

    def expandCmd(self, tour, cmd):
        f = self.expFunctions.get(cmd)
        if f is None:
            return "???" + cmd + "???"
        return f(tour)

    def __enter__(self):
        self.output.write("<Events>\n")
        return self

    def __exit__(self, *args):
        self.output.write("</Events>\n")
        self.output.close()

    def handleTermin(self, termin):
        pass

    def handleTour(self, tour):
        with open(self.xmlFile, "r", encoding="utf-8") as input:
            for l in input:
                mp = paramRE.search(l, 0)
                if mp is not None:
                    sp = mp.span()
                    cmd = l[sp[0]+2:sp[1]-1]
                    l = l[0:sp[0]] + self.expandCmd(tour, cmd) + l[sp[1]:]
                    self.output.writelines([l]);
                elif l.find("<Expand") > 0:
                    ll = self.expandCmd(tour, l.strip())
                    self.output.writelines(ll);
                else:
                    self.output.writelines([l]);

"""
Categories Expansion:
        <Category id="2">
            <i18nName>
                <I18n id="11">
                    <de>
                        <![CDATA[ Bühnenkunst ]]>
                    </de>
                    <en>
                        <![CDATA[ Stage Art ]]>
                    </en>
                </I18n>
            </i18nName>
        </Category>
        <Category id="3">
            <i18nName>
                <I18n id="12">
                    <de>
                        <![CDATA[ Theater ]]>
                    </de>
                    <en>
                        <![CDATA[ Theatre ]]>
                    </en>
                </I18n>
            </i18nName>
        </Category>
        <Category id="8">
            <i18nName>
                <I18n id="17">
                    <de>
                        <![CDATA[ Lesungen ]]>
                    </de>
                    <en>
                        <![CDATA[ Reading ]]>
                    </en>
                </I18n>
            </i18nName>
        </Category>
        
???        
    <criteria>
        <!-- IDs depend upon customer project -->
        <Criterion id="1"/>
        <Criterion id="2"/>
    </criteria>

pricing??    
        <fromPrice>$5.0</fromPrice>
        <toPrice>10.0</toPrice>
        <absolutePrice>15.0</absolutePrice>
        <freeOfCharge>false</freeOfCharge>
        <priceDescription>
            <I18n>
                <de>
                    <![CDATA[priceDescription]]>
                </de>
            </I18n>
        </priceDescription>
        
    <bookingLink>
        <!-- has to be a valid URL -->
        <I18n>
            <de>
                <![CDATA[http://www.booking-link.de]]>
            </de>
            <en>
                <![CDATA[http://www.booking-link.en]]>
            </en>
        </I18n>
    </bookingLink>

https://intern-touren-termine.adfc.de/api/images/2b6a400f-d5ac-46bf-9133-b53ecd5a180c/download
https://intern-touren-termine.adfc.de/api/images/7e0249e0-0d44-42cf-93b5-5a0f3b76a9ed/download
https://intern-touren-termine.adfc.de/api/images/3e53feb6-3ce3-4b21-9c09-6744a02b7c94/download

imageTitle?? imageDescription?
                   <title>
                        <I18n>
                            <de>
                                <![CDATA[${imageTitle}]]>
                            </de>
                        </I18n>
                    </title>
                    <description>
                        <I18n>
                            <de>
                                <![CDATA[image-description]]>
                            </de>
                        </I18n>
                    </description>
imageType?? 
             <imageType>
                <!-- note: depending on customer project -->
                <ImageType id="1"/>
            </imageType>


"""












