interface DiscordEmbed {}
class BoardGameNightsDiscord {
    readonly calendarIcon: string = "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Google_Calendar_icon.svg/64px-Google_Calendar_icon.svg.png"
    readonly calendarId: string

    constructor(calendarId: string) {
        this.calendarId = calendarId
    }
    shareEventsToDiscord() {
        let embeds = []
        for (let event of this.listUpdatedEvents()) {
            let embed = this.prepareEmbedForDiscord(event)
            if (!embed) { continue }
            embeds.push(embed)
        }
        // Keep only the last events to avoid flooding & respect Discord limits
        // Only 1 event expected in nominal cases
        this.pushEmbedsOnDiscord(embeds.slice(-10))
    }
    prepareEmbedForDiscord(event: GoogleAppsScript.Calendar.Schema.Event) {
        if (event.status != 'confirmed') { return }
        let embed: DiscordEmbed = {
            title: event.summary,
            url: event.htmlLink,
            thumbnail: { "url": this.calendarIcon }
        }
        if (event.start && event.start.dateTime) {
            let dt = new Date(event.start.dateTime)
            embed['description'] = `Organisé le ${dt.toLocaleString()}`
        } else if (event.start && event.start.date) {
            let dt = new Date(event.start.date)
            embed['description'] = `Organisé le ${dt.toLocaleDateString()}`
        }
        return embed
    }
    * listUpdatedEvents() {
        // TODO: support pagination
        let syncToken = this.getProperty("syncToken")        
        let res = Calendar.Events.list(this.calendarId, {
            orderBy: (syncToken ? undefined : 'updated'),
            syncToken: syncToken
        })
        Logger.log(`Found ${res.items.length} updated events from last trigger`)
        if (!res.nextSyncToken) {
            throw new Error('Missing nextSyncToken in the response')
        }
        this.setProperty("syncToken", res.nextSyncToken)
        for (let event of res.items) {
            yield event
        }
    }
    pushEmbedsOnDiscord(embeds: DiscordEmbed[]): void {
        if (embeds.length == 0) { return }
        Logger.log(`Push on Discord ${embeds.length} embeds`)
        // POST the msg to the webhook
        let webhookUrl = this.getProperty('webhookUrl')
        let res = UrlFetchApp.fetch(webhookUrl, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({ embeds: embeds })
        })
        // handle HTTP errors
        let statusCode = res.getResponseCode()
        if (statusCode < 200 || 300 <= statusCode) {
            throw new Error(
                `Unexpected HTTP Status Code ${statusCode} ${res.getContentText()}`)
        }
    }
    getProperty(name: string) {
        let props = PropertiesService.getScriptProperties()
        let qname = `${this.calendarId}/${name}`
        return props.getProperty(qname)        
    }
    setProperty(name: string, value: string) {
        let props = PropertiesService.getScriptProperties()
        let qname = `${this.calendarId}/${name}`
        return props.setProperty(qname, value)        
    }
}

function onCalendarUpdate(event) : void {
    // https://developers.google.com/apps-script/guides/triggers/events#google_calendar_events
    Logger.log("Calendar update triggered")
    let bgnd = new BoardGameNightsDiscord(event.calendarId)
    bgnd.shareEventsToDiscord()
}
