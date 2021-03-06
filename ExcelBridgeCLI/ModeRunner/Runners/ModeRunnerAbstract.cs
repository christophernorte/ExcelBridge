﻿using ExcelBridgeApi.Writer;
using ExcelBridgeCli.Argument;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelBridgeCli.ModeRunner.Runners
{
    public abstract class ModeRunnerAbstract
    {
        protected Options options;
        protected WriterProcessor writer;

        public ModeRunnerAbstract()
        {
            this.writer = new WriterProcessor();
        }

        public ModeRunnerAbstract(Options options) : this()
        {
            this.options = options;
        }

        protected string getCellStringPart(string rawCell)
        {
            Regex regex = new Regex(@"[a-zA-Z]+");
            Match match = regex.Match(rawCell);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "";
            }
        }

        protected uint getCellNumericPart(string rawCell)
        {
            Regex regex = new Regex(@"[\d]");
            Match match = regex.Match(rawCell);
            if (match.Success)
            {
                uint value = 0;
                if (uint.TryParse(match.Value, out value))
                {
                    return value;
                }
                return value;
            }
            else
            {
                return 0;
            }
        }

    }
}
