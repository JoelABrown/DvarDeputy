using Microsoft.AspNetCore.Mvc;
using System.Collections.Concurrent;

namespace Mooseware.DvarDeputy.Controllers;

public class HomeController(ConcurrentQueue<ApiMessage> msgQueue) : Controller
{
    /// <summary>
    /// Message queue to pass requests received by this controller to the UI thread for handling
    /// </summary>
    private readonly ConcurrentQueue<ApiMessage> _messageQueue = msgQueue;

    [HttpGet("/status")]
    public ActionResult<string> Status()
    {
        // Let the caller know that the listener is listening
        string status = "alive";
        return Ok(status);
    }

    /// <summary>
    /// HTTP API controller for Viewer controls (e.g. show|hide)
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/viewer")]
    [HttpPost()]
    public IActionResult Viewer(string cmd = "?")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.ViewerShow
         || command == ApiMessage.ViewerHide)
        {
            ApiMessage message = new()
            {
                Verb = ApiMessageVerb.Viewer,
                Parameters = cmd
            };
            _messageQueue.Enqueue(message);
            return Ok();
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Viewer controls (e.g. previous|next|first)
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/page")]
    [HttpPost()]
    public IActionResult Page(string cmd = "?")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.PagePrevious
         || command == ApiMessage.PageNext
         || command == ApiMessage.PageFirst)
        {
            ApiMessage message = new()
            {
                Verb = ApiMessageVerb.Page,
                Parameters = cmd
            };
            _messageQueue.Enqueue(message);
            return Ok();
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Scrolling controls (e.g. (e.g. forward|backward|stop))
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/scroll")]
    [HttpPost()]
    public IActionResult Scroll(string cmd = "?")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.ScrollForward
         || command == ApiMessage.ScrollBackward
         || command == ApiMessage.ScrollStop)
        {
            ApiMessage message = new()
            {
                Verb = ApiMessageVerb.Scroll,
                Parameters = cmd
            };
            _messageQueue.Enqueue(message);
            return Ok();
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Font sizing controls (e.g. (e.g. increase|decrease|reset))
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/font")]
    [HttpPost()]
    public IActionResult Font(string cmd = "?", string val = "0")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.FontIncrease
         || command == ApiMessage.FontDecrease
         || command == ApiMessage.FontReset)
        {
            if (command == ApiMessage.FontIncrease ||  command == ApiMessage.FontDecrease)
            {
                // The scalar value is also of interest. It needs to be a number.
                if (double.TryParse(val, out var numericValue))
                {
                    ApiMessage message = new()
                    {
                        Verb = ApiMessageVerb.Font,
                        Parameters = cmd,
                        Scalar = numericValue
                    };
                    _messageQueue.Enqueue(message);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            else
            {   // Reset.
                ApiMessage message = new()
                {
                    Verb = ApiMessageVerb.Font,
                    Parameters = cmd
                };
                _messageQueue.Enqueue(message);
                return Ok();
            }
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Scroll Speed sizing controls (e.g. (e.g. increase|decrease|reset))
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/speed")]
    [HttpPost()]
    public IActionResult ScrollSpeed(string cmd = "?", string val = "0")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.ScrollSpeedIncrease
         || command == ApiMessage.ScrollSpeedDecrease
         || command == ApiMessage.ScrollSpeedReset)
        {
            if (command == ApiMessage.ScrollSpeedIncrease || command == ApiMessage.ScrollSpeedDecrease)
            {
                // The scalar value is also of interest. It needs to be a number.
                if (double.TryParse(val, out var numericValue))
                {
                    ApiMessage message = new()
                    {
                        Verb = ApiMessageVerb.ScrollSpeed,
                        Parameters = cmd,
                        Scalar = Math.Round(numericValue, 2)    // Only take 2 decimal places.
                    };
                    _messageQueue.Enqueue(message);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            else
            {   // Reset.
                ApiMessage message = new()
                {
                    Verb = ApiMessageVerb.ScrollSpeed,
                    Parameters = cmd
                };
                _messageQueue.Enqueue(message);
                return Ok();
            }
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Spacing controls (e.g. (e.g. increase|decrease|reset))
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/spacing")]
    [HttpPost()]
    public IActionResult Spacing(string cmd = "?", string val = "0")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.SpacingIncrease
         || command == ApiMessage.SpacingDecrease
         || command == ApiMessage.SpacingReset)
        {
            if (command == ApiMessage.SpacingIncrease || command == ApiMessage.SpacingDecrease)
            {
                // The scalar value is also of interest. It needs to be a number.
                if (double.TryParse(val, out var numericValue))
                {
                    ApiMessage message = new()
                    {
                        Verb = ApiMessageVerb.Spacing,
                        Parameters = cmd,
                        Scalar = Math.Round(numericValue, 2)    // Only take 2 decimal places.
                    };
                    _messageQueue.Enqueue(message);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            else
            {   // Reset.
                ApiMessage message = new()
                {
                    Verb = ApiMessageVerb.Spacing,
                    Parameters = cmd
                };
                _messageQueue.Enqueue(message);
                return Ok();
            }
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Margin controls (e.g. (e.g. increase|decrease|reset))
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/margin")]
    [HttpPost()]
    public IActionResult Margin(string cmd = "?", string val = "0")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.MarginIncrease
         || command == ApiMessage.MarginDecrease
         || command == ApiMessage.MarginReset)
        {
            if (command == ApiMessage.MarginIncrease || command == ApiMessage.MarginDecrease)
            {
                // The scalar value is also of interest. It needs to be a number.
                if (double.TryParse(val, out var numericValue))
                {
                    ApiMessage message = new()
                    {
                        Verb = ApiMessageVerb.Margin,
                        Parameters = cmd,
                        Scalar = Math.Round(numericValue, 2)    // Only take 2 decimal places.
                    };
                    _messageQueue.Enqueue(message);
                    return Ok();
                }
                else
                {
                    return BadRequest();
                }
            }
            else
            {   // Reset.
                ApiMessage message = new()
                {
                    Verb = ApiMessageVerb.Margin,
                    Parameters = cmd
                };
                _messageQueue.Enqueue(message);
                return Ok();
            }
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Theme controls (e.g. light|dark)
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/theme")]
    [HttpPost()]
    public IActionResult Theme(string cmd = "?")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.ThemeLight
         || command == ApiMessage.ThemeDark
         || command == ApiMessage.ThemeMatrix)
        {
            ApiMessage message = new()
            {
                Verb = ApiMessageVerb.Theme,
                Parameters = cmd
            };
            _messageQueue.Enqueue(message);
            return Ok();
        }
        else
        {
            return BadRequest();
        }
    }

    /// <summary>
    /// HTTP API controller for Progress Bug controls (e.g. none|side|bottom)
    /// </summary>
    /// <param name="cmd">A string indicating the type of action being requested</param>
    /// <returns>HTTP response code (200 if OK)</returns>
    [Route("/bug")]
    [HttpPost()]
    public IActionResult Bug(string cmd = "?")
    {
        // Verify that the command is acceptable
        string command = cmd.ToLower().Trim();
        if (command == ApiMessage.BugNone
         || command == ApiMessage.BugSide
         || command == ApiMessage.BugBottom)
        {
            ApiMessage message = new()
            {
                Verb = ApiMessageVerb.Bug,
                Parameters = cmd
            };
            _messageQueue.Enqueue(message);
            return Ok();
        }
        else
        {
            return BadRequest();
        }
    }
}
