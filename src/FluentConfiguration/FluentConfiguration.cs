// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Represents the fluent configuration for the specfidied model.
    /// </summary>
    /// <typeparam name="TModel">The type of model.</typeparam>
    public class FluentConfiguration<TModel> : IFluentConfiguration where TModel : class
    {
        private IDictionary<string, PropertyConfiguration> _propertyConfigs;
        private IList<PropertyInfo> _propertyInfo;
        private IList<StatisticsConfig> _statisticsConfigs;
        private IList<FilterConfig> _filterConfigs;
        private IList<FreezeConfig> _freezeConfigs;

        /// <summary>
        /// Initializes a new instance of the <see cref="FluentConfiguration{TModel}"/> class.
        /// </summary>
        public FluentConfiguration()
        {
            var mType = typeof(TModel);
            _propertyInfo = mType.GetProperties();

            _propertyConfigs = new Dictionary<string, PropertyConfiguration>();
            _statisticsConfigs = new List<StatisticsConfig>();
            _filterConfigs = new List<FilterConfig>();
            _freezeConfigs = new List<FreezeConfig>();
        }

        /// <summary>
        /// Gets the property configs.
        /// </summary>
        /// <value>The property configs.</value>
        IDictionary<string, PropertyConfiguration> IFluentConfiguration.PropertyConfigs
        {
            get
            {
                return _propertyConfigs;
            }
        }

        /// <summary>
        /// Gets the statistics configs.
        /// </summary>
        /// <value>The statistics config.</value>
        IList<StatisticsConfig> IFluentConfiguration.StatisticsConfigs
        {
            get
            {
                return _statisticsConfigs;
            }
        }

        /// <summary>
        /// Gets the filter configs.
        /// </summary>
        /// <value>The filter config.</value>
        IList<FilterConfig> IFluentConfiguration.FilterConfigs
        {
            get
            {
                return _filterConfigs;
            }
        }

        /// <summary>
        /// Gets the freeze configs.
        /// </summary>
        /// <value>The freeze config.</value>
        IList<FreezeConfig> IFluentConfiguration.FreezeConfigs
        {
            get
            {
                return _freezeConfigs;
            }
        }

        /// <summary>
        /// Gets the property configuration by the specified property expression for the specified <typeparamref name="TModel"/> and its <typeparamref name="TProperty"/>.
        /// </summary>
        /// <returns>The <see cref="PropertyConfiguration"/>.</returns>
        /// <param name="propertyExpression">The property expression.</param>
        /// <typeparam name="TProperty">The type of parameter.</typeparam>
        public PropertyConfiguration Property<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            var propertyInfo = GetPropertyInfo(propertyExpression);

            PropertyConfiguration pc = null;
            if (!_propertyConfigs.TryGetValue(propertyInfo.Name, out pc))
            {
                pc = new PropertyConfiguration();
                _propertyConfigs[propertyInfo.Name] = pc;
            }

            return pc;
        }

        /// <summary>
        /// From Annotations
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="fc"></param>
        public FluentConfiguration<TModel> SetIgnore(params Expression<Func<TModel, object>>[] propertyExpressions)
        {
            if (propertyExpressions == null)
                return this;

            foreach (var propertyExpression in propertyExpressions)
            {
                var propertyInfo = GetPropertyInfo(propertyExpression);
                PropertyConfiguration pc = null;
                if (_propertyConfigs.TryGetValue(propertyInfo.Name, out pc))
                    pc.IsIgnored(true, true);
            }

            return this;
        }

        /// <summary>
        /// Auto Gen Index
        /// </summary>
        public void AutoIndex()
        {
            var config = _propertyConfigs;
            var mType = typeof(TModel);
            var index = 0;

            PropertyConfiguration pc = null;
            CellConfig cc = null;
            foreach (var prop in _propertyInfo)
            {
                if (!_propertyConfigs.TryGetValue(prop.Name, out pc))
                    continue;

                cc = pc.CellConfig;
                if (cc.IsExportIgnored || !cc.AutoIndex)
                    continue;

                while (_propertyConfigs.Values.Any(x => x.CellConfig.Index == index))
                {
                    index++;
                }

                pc.HasExcelIndex(index);
                index++;
            }
        }

        /// <summary>
        /// From Annotations
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="fc"></param>
        public FluentConfiguration<TModel> FromAnnotations()
        {
            var pConfig = this._propertyConfigs;
            foreach (var prop in _propertyInfo)
            {
                SetPropertyFromDisplay(prop, pConfig);
            }
            return this;
        }

        /// <summary>
        /// Set Property From DisplayAttribute
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        private static PropertyConfiguration SetPropertyFromDisplay(PropertyInfo propertyInfo,
            IDictionary<string, PropertyConfiguration> _propertyConfigs)
        {
            var display = propertyInfo.GetCustomAttribute<DisplayAttribute>();
            var displayFormat = propertyInfo.GetCustomAttribute<DisplayFormatAttribute>();
            PropertyConfiguration pc = null;
            if (!_propertyConfigs.TryGetValue(propertyInfo.Name, out pc))
            {
                pc = new PropertyConfiguration();
                _propertyConfigs[propertyInfo.Name] = pc;
            }

            if (display != null)
            {
                pc.HasExcelTitle(display.Name);
                if (display.GetOrder().HasValue)
                    pc.HasExcelIndex(display.Order);
            }
            else
            {
                pc.HasExcelTitle(propertyInfo.Name);
            }

            if (displayFormat != null)
            {
                pc.HasDataFormatter(displayFormat.DataFormatString
                                                 .Replace("{0:", "")
                                                 .Replace("}", ""));
            }

            if (pc.CellConfig.Index < 0)
                pc.HasAutoIndex();

            return pc;
        }

        /// <summary>
        /// Configures the statistics for the specified <typeparamref name="TModel"/>. Only for vertical, not for horizontal statistics.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="name">The statistics name. (e.g. Total). In current version, the default name location is (last row, first cell)</param>
        /// <param name="formula">The cell formula, such as SUM, AVERAGE and so on, which applyable for vertical statistics..</param>
        /// <param name="columnIndexes">The column indexes for statistics. if <paramref name="formula"/>is SUM, and <paramref name="columnIndexes"/> is [1,3],
        /// for example, the column No. 1 and 3 will be SUM for first row to last row.</param>
        public FluentConfiguration<TModel> HasStatistics(string name, string formula, params int[] columnIndexes)
        {
            var statistics = new StatisticsConfig
            {
                Name = name,
                Formula = formula,
                Columns = columnIndexes,
            };

            _statisticsConfigs.Add(statistics);

            return this;
        }

        /// <summary>
        /// Configures the excel filter behaviors for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="firstColumn">The first column index.</param>
        /// <param name="lastColumn">The last column index.</param>
        /// <param name="firstRow">The first row index.</param>
        /// <param name="lastRow">The last row index. If is null, the value is dynamic calculate by code.</param>
        public FluentConfiguration<TModel> HasFilter(int firstColumn, int lastColumn, int firstRow, int? lastRow = null)
        {
            var filter = new FilterConfig
            {
                FirstCol = firstColumn,
                FirstRow = firstRow,
                LastCol = lastColumn,
                LastRow = lastRow,
            };

            _filterConfigs.Add(filter);

            return this;
        }

        /// <summary>
        /// Configures the excel freeze behaviors for the specified <typeparamref name="TModel"/>.
        /// </summary>
        /// <returns>The <see cref="FluentConfiguration{TModel}"/>.</returns>
        /// <param name="columnSplit">The column number to split.</param>
        /// <param name="rowSplit">The row number to split.param>
        /// <param name="leftMostColumn">The left most culomn index.</param>
        /// <param name="topMostRow">The top most row index.</param>
        public FluentConfiguration<TModel> HasFreeze(int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
        {
            var freeze = new FreezeConfig
            {
                ColSplit = columnSplit,
                RowSplit = rowSplit,
                LeftMostColumn = leftMostColumn,
                TopRow = topMostRow,
            };

            _freezeConfigs.Add(freeze);

            return this;
        }

        private PropertyInfo GetPropertyInfo<TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            if (propertyExpression.NodeType != ExpressionType.Lambda)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            var lambda = (LambdaExpression)propertyExpression;

            var memberExpression = ExtractMemberExpression(lambda.Body);
            if (memberExpression == null)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            if (memberExpression.Member.DeclaringType == null)
            {
                throw new InvalidOperationException("Property does not have declaring type");
            }

            return memberExpression.Member.DeclaringType.GetProperty(memberExpression.Member.Name);
        }

        private MemberExpression ExtractMemberExpression(Expression expression)
        {
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                return ((MemberExpression)expression);
            }

            if (expression.NodeType == ExpressionType.Convert)
            {
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            }

            return null;
        }
    }
}