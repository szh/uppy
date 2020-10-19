const Provider = require('../Provider')

const request = require('request')
const purest = require('purest')({ request })
const logger = require('../../logger')
const adapter = require('./adapter')
const { ProviderApiError, ProviderAuthError } = require('../error')

/**
 * Adapter for API https://docs.microsoft.com/en-us/onedrive/developer/rest-api/
 */
class OneDrive extends Provider {
  constructor (options) {
    super(options)
    this.authProvider = options.provider = OneDrive.authProvider
    this.client = purest(options)
  }

  static get authProvider () {
    return 'microsoft'
  }

  _userInfo ({ token }, done) {
    this.client
      .get('me')
      .auth(token)
      .request(done)
  }

  /**
   * Makes 2 requests in parallel - 1. to get files, 2. to get user email
   * it then waits till both requests are done before proceeding with the callback
   *
   * @param {object} options
   * @param {function} done
   */
  list ({ directory, query, token }, done) {
    let path
    const qs = {}
    if (!query.driveId) {
      path = '/me/drives'
    } else if (query.driveId === '_listsites_') {
      path = '/sites?search='
    } else {
      path = `/drives/${query.driveId}/`
      if (!!directory && directory !== 'root') {
        path += `items/${directory}`
      } else {
        path += 'root'
      }
      path += '/children'
      qs.$expand = 'thumbnails'
    }

    if (query.cursor) {
      qs.$skiptoken = query.cursor
    }

    console.log('path: ' + path)

    this.client
      .get(path)
      .qs(qs)
      .auth(token)
      .request((err, resp, body) => {
        if (err || resp.statusCode !== 200) {
          err = this._error(err, resp)
          logger.error(err, 'provider.onedrive.list.error')
          return done(err)
        } else {
          this._userInfo({ token }, (err, infoResp) => {
            if (err || infoResp.statusCode !== 200) {
              err = this._error(err, infoResp)
              logger.error(err, 'provider.onedrive.user.error')
              return done(err)
            }
            console.log('body: ' + JSON.stringify(body))
            if (query.driveId === '_listsites_') {
              this.adaptSharepointSitesData(body, infoResp.body.mail || infoResp.body.userPrincipalName, token, done)
            } else {
              done(null, this.adaptData(body, infoResp.body.mail || infoResp.body.userPrincipalName, !query.driveId))
            }
          })
        }
      })
  }

  download ({ id, token, query }, onData) {
    const rootPath = query.driveId ? `/drives/${query.driveId}` : '/me/drive'
    return this.client
      .get(`${rootPath}/items/${id}/content`)
      .auth(token)
      .request()
      .on('response', (resp) => {
        if (resp.statusCode !== 200) {
          onData(this._error(null, resp))
        } else {
          resp.on('data', (chunk) => onData(null, chunk))
        }
      })
      .on('end', () => onData(null, null))
      .on('error', (err) => {
        logger.error(err, 'provider.onedrive.download.error')
        onData(err)
      })
  }

  thumbnail (_, done) {
    // not implementing this because a public thumbnail from onedrive will be used instead
    const err = new Error('call to thumbnail is not implemented')
    logger.error(err, 'provider.onedrive.thumbnail.error')
    return done(err)
  }

  size ({ id, query, token }, done) {
    const rootPath = query.driveId ? `/drives/${query.driveId}` : '/me/drive'
    return this.client
      .get(`${rootPath}/items/${id}`)
      .auth(token)
      .request((err, resp, body) => {
        if (err || resp.statusCode !== 200) {
          err = this._error(err, resp)
          logger.error(err, 'provider.onedrive.size.error')
          return done(err)
        } else {
          done(null, body.size)
        }
      })
  }

  logout (_, done) {
    // access revoke is not supported by Microsoft/OneDrive's API
    done(null, { revoked: false, manual_revoke_url: 'https://account.live.com/consent/Manage' })
  }

  adaptData (res, username, includeRemoteFolder) {
    const data = { username, items: [] }
    const items = adapter.getItemSubList(res)
    items.forEach((item) => {
      data.items.push({
        isFolder: adapter.isFolder(item),
        icon: adapter.getItemIcon(item),
        name: adapter.getItemName(item),
        mimeType: adapter.getMimeType(item),
        id: adapter.getItemId(item),
        thumbnail: adapter.getItemThumbnailUrl(item),
        requestPath: adapter.getItemRequestPath(item),
        modifiedDate: adapter.getItemModifiedDate(item),
        size: adapter.getItemSize(item)
      })
    })

    if (includeRemoteFolder) {
      // Add a pseudo-folder for listing sharepoint sites
      data.items.push({
        isFolder: true,
        icon: 'folder',
        name: 'Other Remote Drives',
        id: 'root',
        requestPath: 'root?driveId=_listsites_'
      })
    }

    data.nextPagePath = adapter.getNextPagePath(res)

    return data
  }

  adaptSharepointSitesData (res, username, token, done) {
    const items = adapter.getItemSubList(res)
    let loadedSites = 0
    const data = { username, items: [] }
    items.forEach((item) => {
      const siteName = item.displayName
      this.client
        .get('sites/' + item.id + '/drives')
        .auth(token)
        .request((err, resp, body) => {
          if (err || resp.statusCode !== 200) {
            err = this._error(err, resp)
            logger.error(err, 'provider.onedrive.list.error')
            return done(err)
          } else {
            console.log('body: ' + JSON.stringify(body))
            const siteDrives = this.adaptData(body).items
            siteDrives.forEach((item) => {
              item.name = siteName + ' ' + item.name
            })
            loadedSites++
            data.items.push(siteDrives)

            if (loadedSites === items.length) {
              data.nextPagePath = adapter.getNextPagePath(res)
              done(null, data)
            }
          }
        })
    })
  }

  _error (err, resp) {
    if (resp) {
      const fallbackMsg = `request to ${this.authProvider} returned ${resp.statusCode}`
      const errMsg = (resp.body || {}).error ? resp.body.error.message : fallbackMsg
      return resp.statusCode === 401 ? new ProviderAuthError() : new ProviderApiError(errMsg, resp.statusCode)
    }

    return err
  }
}

module.exports = OneDrive
