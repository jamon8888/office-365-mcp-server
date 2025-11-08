/**
 * JSON Batching Manager for Microsoft Graph API
 * Enables efficient batch processing of multiple API requests
 * Supports up to 20 requests per batch with automatic chunking
 */

const { callGraphAPI } = require('./graph-api');

class BatchRequestManager {
  constructor() {
    this.maxBatchSize = 20; // Microsoft Graph API limit
  }

  /**
   * Execute batch requests with automatic chunking and optional parallel execution
   * @param {Array} requests - Array of request objects
   * @param {string} accessToken - Access token for authentication
   * @param {boolean} parallel - Execute chunks in parallel (default: false for safety)
   * @returns {Array} Array of responses from all batches
   */
  async executeBatch(requests, accessToken, parallel = false) {
    if (!requests || requests.length === 0) {
      return [];
    }

    const chunks = this.chunkRequests(requests);
    
    if (parallel && chunks.length > 1) {
      // Execute chunks in parallel for better performance
      const results = await Promise.allSettled(
        chunks.map(chunk => this.executeSingleBatch(chunk, accessToken))
      );
      
      // Flatten results and handle errors
      return results.flatMap(result => {
        if (result.status === 'fulfilled') {
          return result.value;
        } else {
          console.error('Parallel batch request failed:', result.reason);
          return [];
        }
      });
    } else {
      // Sequential execution (safer default)
      const results = [];
      for (const chunk of chunks) {
        const chunkResults = await this.executeSingleBatch(chunk, accessToken);
        results.push(...chunkResults);
      }
      return results;
    }
  }

  /**
   * Execute a single batch request
   * @param {Array} chunk - Array of requests for this batch
   * @param {string} accessToken - Access token for authentication
   * @returns {Array} Array of responses from the batch
   */
  async executeSingleBatch(chunk, accessToken) {
    try {
      const batchResponse = await callGraphAPI(
        accessToken,
        'POST',
        '$batch',
        { requests: chunk }
      );
      
      // Extract responses from batch response
      if (batchResponse && batchResponse.responses) {
        return batchResponse.responses;
      }
      return [];
    } catch (error) {
      console.error('Batch request failed:', error);
      // Add error responses for failed batch
      return chunk.map(req => ({
        id: req.id,
        status: 500,
        body: { error: { message: 'Batch request failed', details: error.message } }
      }));
    }
  }

  /**
   * Split requests into chunks of maximum batch size
   * @param {Array} requests - Array of requests to chunk
   * @returns {Array} Array of request chunks
   */
  chunkRequests(requests) {
    const chunks = [];
    for (let i = 0; i < requests.length; i += this.maxBatchSize) {
      chunks.push(requests.slice(i, i + this.maxBatchSize));
    }
    return chunks;
  }

  /**
   * Create a properly formatted batch request object
   * @param {string} id - Unique identifier for the request
   * @param {string} method - HTTP method (GET, POST, PATCH, DELETE)
   * @param {string} url - Relative URL for the API endpoint
   * @param {Object} body - Request body (optional)
   * @param {Object} headers - Additional headers (optional)
   * @returns {Object} Formatted batch request object
   */
  createBatchRequest(id, method, url, body = null, headers = {}) {
    const request = {
      id: id.toString(),
      method: method.toUpperCase(),
      url: url.startsWith('/') ? url : `/${url}`,
      headers: headers
    };
    
    if (body && (method === 'POST' || method === 'PATCH' || method === 'PUT')) {
      request.body = body;
    }
    
    return request;
  }

  /**
   * Helper method to batch mark emails as read/unread
   * @param {Array} emailIds - Array of email IDs
   * @param {boolean} isRead - Mark as read (true) or unread (false)
   * @param {string} accessToken - Access token
   * @returns {Object} Summary of batch operation results
   */
  async batchUpdateEmails(emailIds, isRead, accessToken) {
    const requests = emailIds.map((id, index) => 
      this.createBatchRequest(
        index.toString(),
        'PATCH',
        `/me/messages/${id}`,
        { isRead: isRead }
      )
    );
    
    const results = await this.executeBatch(requests, accessToken);
    
    return {
      total: emailIds.length,
      successful: results.filter(r => r.status >= 200 && r.status < 300).length,
      failed: results.filter(r => r.status >= 400).length,
      results: results
    };
  }

  /**
   * Helper method to batch move emails to a folder
   * @param {Array} emailIds - Array of email IDs
   * @param {string} destinationFolderId - Target folder ID
   * @param {string} accessToken - Access token
   * @returns {Object} Summary of batch operation results
   */
  async batchMoveEmails(emailIds, destinationFolderId, accessToken) {
    const requests = emailIds.map((id, index) => 
      this.createBatchRequest(
        index.toString(),
        'POST',
        `/me/messages/${id}/move`,
        { destinationId: destinationFolderId }
      )
    );
    
    const results = await this.executeBatch(requests, accessToken);
    
    return {
      total: emailIds.length,
      successful: results.filter(r => r.status >= 200 && r.status < 300).length,
      failed: results.filter(r => r.status >= 400).length,
      results: results
    };
  }

  /**
   * Helper method to batch delete emails
   * @param {Array} emailIds - Array of email IDs
   * @param {string} accessToken - Access token
   * @returns {Object} Summary of batch operation results
   */
  async batchDeleteEmails(emailIds, accessToken) {
    const requests = emailIds.map((id, index) => 
      this.createBatchRequest(
        index.toString(),
        'DELETE',
        `/me/messages/${id}`
      )
    );
    
    const results = await this.executeBatch(requests, accessToken);
    
    return {
      total: emailIds.length,
      successful: results.filter(r => r.status >= 200 && r.status < 300).length,
      failed: results.filter(r => r.status >= 400).length,
      results: results
    };
  }

  /**
   * Helper method to batch create calendar events
   * @param {Array} events - Array of event objects
   * @param {string} accessToken - Access token
   * @returns {Object} Summary of batch operation results
   */
  async batchCreateEvents(events, accessToken) {
    const requests = events.map((event, index) => 
      this.createBatchRequest(
        index.toString(),
        'POST',
        '/me/calendar/events',
        event
      )
    );
    
    const results = await this.executeBatch(requests, accessToken);
    
    return {
      total: events.length,
      successful: results.filter(r => r.status >= 200 && r.status < 300).length,
      failed: results.filter(r => r.status >= 400).length,
      results: results
    };
  }

  /**
   * Parse batch responses and extract successful results
   * @param {Array} responses - Array of batch responses
   * @returns {Array} Array of successful response bodies
   */
  parseSuccessfulResponses(responses) {
    return responses
      .filter(r => r.status >= 200 && r.status < 300)
      .map(r => r.body);
  }

  /**
   * Parse batch responses and extract errors
   * @param {Array} responses - Array of batch responses
   * @returns {Array} Array of error details
   */
  parseErrorResponses(responses) {
    return responses
      .filter(r => r.status >= 400)
      .map(r => ({
        id: r.id,
        status: r.status,
        error: r.body?.error || { message: 'Unknown error' }
      }));
  }
}

// Export singleton instance
module.exports = new BatchRequestManager();