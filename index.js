'use strict';

const axios = require('axios')
const logger = require('winston').createLogger({})
const _ = require('underscore')
const mime = require('mime-types')
const fs = require('fs')

const tenantID = ''
const oAuthClientID = ''
const clientSecret = ''
if (!tenantID || !oAuthClientID || !clientSecret) {
    throw new Error("Credentials missing for 'microsoft-graph authentication.")
}

module.exports = {
    name: 'microsoft-graph',
    version: '0.1.0',
    send: (mail, callback) => {

        Promise.all(
            _.map(
                mail.data.attachments ? mail.data.attachments : [],
                (x) => fs.promises.readFile(x.path, 'base64').then(bytes => {
                    return {
                        '@odata.type': '#microsoft.graph.fileAttachment',
                        name: x.filename,
                        contentType: mime.lookup(x.filename),
                        contentId: x.cid,
                        contentBytes: bytes
                    }
                })
            )
        ).then(attachments => axios({
                method: 'post',
                url: `https://login.microsoftonline.com/${tenantID}/oauth2/v2.0/token`,
                data: new URLSearchParams({
                    client_id: oAuthClientID,
                    client_secret: clientSecret,
                    scope: "https://graph.microsoft.com/.default",
                    grant_type: "client_credentials"
                }).toString()
            }).then(r => {

                // Make call to send email.
                // Documentation can be found at:
                // https://learn.microsoft.com/en-us/graph/api/user-sendmail
                axios({
                    method: 'post',
                    url: `https://graph.microsoft.com/v1.0/users/${mail.data.from}/sendMail`,
                    headers: {
                        'Authorization': "Bearer " + r.data.access_token,
                        'Content-Type': 'application/json'
                    },
                    data: {
                        // Documentation can be found at:
                        // https://learn.microsoft.com/en-us/graph/api/resources/message
                        message: {
                            subject: mail.data.subject,
                            body: {
                                contentType: 'HTML',
                                content: mail.data.html
                            },
                            toRecipients: _.map(mail.data.to, (x) => {
                                return { emailAddress: {address: x} }
                            }),
                            ccRecipients: _.map(mail.data.cc, (x) => {
                                return { emailAddress: {address: x} }
                            }),
                            bccRecipients: _.map(mail.data.bcc, (x) => {
                                return { emailAddress: {address: x} }
                            }),
                            hasAttachments: mail.data.attachments.length > 0,
                            attachments: attachments
                        },
                        saveToSentItems: false
                    }
                }).then(() => {
                    // Call the callback on success.
                    callback(null, true)
                }).catch(err => {
                    logger.error('Could not send email through Microsoft Graph.', err)
                    callback(null, false)
                })

            }).catch(err => {
                logger.error('Could not authenticate through Microsoft Graph.', err)
                callback(null, false)
            })
        ).catch(err => {
            logger.error('Could not read attachment files.', err)
            callback(null, false)
        })
    }
}
