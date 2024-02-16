/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;

namespace HtmlToOpenXml
{
    using TagsAtSameLevel = System.ArraySegment<DocumentFormat.OpenXml.OpenXmlElement>;

    /// <summary>
    /// Defines the styles to apply on OpenXml elements.
    /// </summary>
    abstract class OpenXmlStyleCollectionBase
    {
        /// <summary>Holds the tags to apply to the current OpenXml element.</summary>
        /// <remarks>The key contains the name of the tag, the values contains a list of queued attributes of the same tag.</remarks>
        protected readonly Dictionary<String, Stack<TagsAtSameLevel>> tags;

        protected OpenXmlStyleCollectionBase()
        {
            tags = new (StringComparer.OrdinalIgnoreCase);
        }

        internal virtual void Reset()
        {
            tags.Clear();
        }

        //____________________________________________________________________
        //

        // Related to tags behaviors: as the tags can be embedded, we need to know which style
        // we should apply on a specific run.
        // Let's take this example:  <font size=3>A<font size=4><strong> big</strong></font> leopard.</font>
        // You see, "big" should be size=4 and not 3. But leopard has its size to 3.

        #region ApplyTags

        /// <summary>
        /// Apply all the current Html tag (Run properties) to the specified run.
        /// </summary>
        public abstract void ApplyTags(OpenXmlCompositeElement element);

        #endregion

        #region BeginTag

        /// <summary>
        /// Add the specified tag to the list.
        /// </summary>
        /// <param name="name">The name of the tag.</param>
        /// <param name="elements">The Run properties to apply to the next build run until the tag is popped out.</param>
        public void BeginTag(string name, List<OpenXmlElement> elements)
        {
            if (elements.Count == 0) return;

            if (!tags.TryGetValue(name, out var enqueuedTags))
            {
                tags.Add(name, enqueuedTags = new Stack<TagsAtSameLevel>());
            }

            enqueuedTags.Push(new TagsAtSameLevel(elements.ToArray()));
        }

        /// <summary>
        /// Add the specified tag to the list.
        /// </summary>
        /// <param name="name">The name of the tag.</param>
        /// <param name="elements">The Run properties to apply to the next build run until the tag is popped out.</param>
        public void BeginTag(string name, params OpenXmlElement[] elements)
        {
            if (!tags.TryGetValue(name, out var enqueuedTags))
            {
                tags.Add(name, enqueuedTags = new Stack<TagsAtSameLevel>());
            }

            enqueuedTags.Push(new TagsAtSameLevel(elements));
        }

        #endregion

        #region MergeTag

        /// <summary>
        /// Merge the properties with the tag of the previous level.
        /// </summary>
        /// <param name="name">The name of the tag.</param>
        /// <param name="elements">The properties to apply to the next build run until the tag is popped out.</param>
        public void MergeTag(string name, List<OpenXmlElement> elements)
        {
            if (!tags.TryGetValue(name, out var enqueuedTags))
            {
                BeginTag(name, elements.ToArray());
            }
            else
            {
                var knonwTags = new Dictionary<string, OpenXmlElement>();
                for (int i = 0; i < elements.Count; i++)
                    if (!knonwTags.ContainsKey(elements[i].LocalName))
                        knonwTags.Add(elements[i].LocalName, elements[i]);

                OpenXmlElement[] array;
                foreach (TagsAtSameLevel tagOfSameLevel in enqueuedTags)
                {
                    array = tagOfSameLevel.Array!;
                    for (int i = 0; i < array.Length; i++)
                    {
                        if (!knonwTags.ContainsKey(array[i].LocalName))
                            knonwTags.Add(array[i].LocalName, array[i]);
                    }
                }

                array = new OpenXmlElement[knonwTags.Count];
                knonwTags.Values.CopyTo(array, 0);
                enqueuedTags.Push(new TagsAtSameLevel(array));
            }
        }

        #endregion

        #region EndTag

        /// <summary>
        /// Remove the specified tag from the list.
        /// </summary>
        /// <param name="name">The name of the tag.</param>
        public virtual void EndTag(string name)
        {
            if (tags.TryGetValue(name, out var enqueuedTags))
            {
                enqueuedTags.Pop();
                if (enqueuedTags.Count == 0) tags.Remove(name);
            }
        }

        #endregion
    }
}